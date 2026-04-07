import { Command } from "commander";
import { graph, validateId, GraphError } from "../lib/graph.js";
import { outputData, handleCommandError } from "../lib/output.js";
import { mkdir, readFile, writeFile, rename, unlink } from "node:fs/promises";
import { homedir } from "node:os";
import { dirname, join } from "node:path";

const WELL_KNOWN_FOLDERS = new Set([
  "inbox",
  "sentitems",
  "drafts",
  "deleteditems",
  "junkemail",
  "archive",
  "outbox",
]);

function resolveFolder(folder: string): string {
  const lower = folder.toLowerCase();
  if (WELL_KNOWN_FOLDERS.has(lower)) return lower;
  validateId(folder, "folder ID");
  return folder;
}

function parseEmailCsv(csv: string, fieldName: string): string[] {
  return csv.split(",").map((e) => {
    const trimmed = e.trim();
    if (!trimmed.includes("@")) {
      throw new Error(`Invalid email address in ${fieldName}: "${trimmed}"`);
    }
    return trimmed;
  });
}

// -------- mail delta helpers --------

const DEFAULT_DELTA_SELECT =
  "internetMessageId,from,subject,receivedDateTime,isRead";

// Graph error codes that indicate a delta token has expired or been invalidated.
const DELTA_EXPIRED_CODES = new Set([
  "syncStateNotFound",
  "syncStateInvalid",
  "ErrorInvalidSyncStateData",
  "resyncRequired",
]);

function getDefaultStateFile(folder: string): string {
  const xdgState = process.env["XDG_STATE_HOME"];
  const base = xdgState ?? join(homedir(), ".local", "state");
  return join(base, "m365-cli", `mail-delta-${folder}.link`);
}

function ensureInternetMessageId(select: string): string {
  const fields = select
    .split(",")
    .map((f) => f.trim())
    .filter(Boolean);
  if (!fields.some((f) => f.toLowerCase() === "internetmessageid")) {
    fields.push("internetMessageId");
  }
  return fields.join(",");
}

async function atomicWrite(filePath: string, content: string): Promise<void> {
  await mkdir(dirname(filePath), { recursive: true });
  const tmp = filePath + ".tmp";
  await writeFile(tmp, content, "utf-8");
  await rename(tmp, filePath);
}

async function readStateFile(filePath: string): Promise<string | null> {
  try {
    const content = await readFile(filePath, "utf-8");
    return content.trim() || null;
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === "ENOENT") return null;
    throw e;
  }
}

async function deleteStateFile(filePath: string): Promise<void> {
  try {
    await unlink(filePath);
  } catch {
    // ignore ENOENT
  }
}

function isDeltaExpired(err: unknown): err is GraphError {
  if (err instanceof GraphError) {
    if (err.status === 410) return true;
    if (DELTA_EXPIRED_CODES.has(err.code)) return true;
  }
  return false;
}

interface DeltaPage {
  value: unknown[];
  "@odata.nextLink"?: string;
  "@odata.deltaLink"?: string;
}

function buildDeltaUrl(
  folder: string,
  opts: { select?: string; changeType?: string }
): string {
  const select = ensureInternetMessageId(opts.select ?? DEFAULT_DELTA_SELECT);
  const params = new URLSearchParams();
  params.set("$select", select);
  if (opts.changeType) params.set("changeType", opts.changeType);
  return `/me/mailFolders/${folder}/messages/delta?${params.toString()}`;
}

async function drainPages(
  startUrl: string,
  stateFile: string,
  quiet: boolean,
  isNdjson: boolean,
  maxPageSize: number
): Promise<void> {
  const preferHeader = { Prefer: `odata.maxpagesize=${maxPageSize}` };
  const allItems: unknown[] = [];
  let url = startUrl;

  while (true) {
    const page = await graph.get<DeltaPage>(url, { headers: preferHeader });

    if (!quiet) {
      for (const item of page.value) {
        if (isNdjson) {
          process.stdout.write(JSON.stringify(item) + "\n");
        } else {
          allItems.push(item);
        }
      }
    }

    if (page["@odata.deltaLink"]) {
      await atomicWrite(stateFile, page["@odata.deltaLink"]);
      break;
    } else if (page["@odata.nextLink"]) {
      url = page["@odata.nextLink"];
    } else {
      break;
    }
  }

  if (!quiet && !isNdjson) {
    outputData(allItems);
  }
}

// -------- end mail delta helpers --------

export function registerMailCommands(program: Command): void {
  const mail = program.command("mail").description("Mail operations");

  // m365 mail list
  mail
    .command("list")
    .description("List messages in a mail folder")
    .option("--folder <name|id>", "Folder name or ID (default: inbox)")
    .option("--filter <odata>", "OData filter expression")
    .option("--select <fields>", "Comma-separated fields to include")
    .option("--top <n>", "Max number of messages to return", "25")
    .action(async (opts) => {
      try {
        const folder = resolveFolder(opts.folder ?? "inbox");
        const params = new URLSearchParams();
        params.set("$top", opts.top);
        if (opts.filter) params.set("$filter", opts.filter);
        if (opts.select) params.set("$select", opts.select);
        const result = await graph.get<{ value: unknown[] }>(
          `/me/mailFolders/${folder}/messages?${params.toString()}`
        );
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 mail get <messageId>
  mail
    .command("get <messageId>")
    .description("Get a message by ID")
    .action(async (messageId) => {
      try {
        validateId(messageId, "message ID");
        const result = await graph.get<unknown>(`/me/messages/${messageId}`);
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 mail send
  mail
    .command("send")
    .description("Send an email message")
    .requiredOption("--to <emails>", "Recipient(s), comma-separated")
    .requiredOption("--subject <str>", "Message subject")
    .requiredOption("--body <str>", "Message body")
    .option("--cc <emails>", "CC recipient(s), comma-separated")
    .option("--bcc <emails>", "BCC recipient(s), comma-separated")
    .option("--importance <level>", "Importance: low, normal, high")
    .option("--html", "Treat body as HTML (default: plain text)")
    .action(async (opts) => {
      try {
        const toAddresses = parseEmailCsv(opts.to, "--to");
        const message: Record<string, unknown> = {
          subject: opts.subject,
          body: {
            contentType: opts.html ? "HTML" : "Text",
            content: opts.body,
          },
          toRecipients: toAddresses.map((a) => ({ emailAddress: { address: a } })),
        };
        if (opts.cc) {
          message.ccRecipients = parseEmailCsv(opts.cc, "--cc").map((a) => ({
            emailAddress: { address: a },
          }));
        }
        if (opts.bcc) {
          message.bccRecipients = parseEmailCsv(opts.bcc, "--bcc").map((a) => ({
            emailAddress: { address: a },
          }));
        }
        if (opts.importance) {
          message.importance = opts.importance;
        }
        await graph.post<void>("/me/sendMail", { message, saveToSentItems: true });
        outputData({ message: "Mail sent." });
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 mail reply <messageId>
  mail
    .command("reply <messageId>")
    .description("Reply to a message")
    .requiredOption("--body <str>", "Reply body")
    .option("--reply-all", "Reply to all recipients")
    .option("--html", "Treat body as HTML (default: plain text)")
    .action(async (messageId, opts) => {
      try {
        validateId(messageId, "message ID");
        const endpoint = opts.replyAll
          ? `/me/messages/${messageId}/replyAll`
          : `/me/messages/${messageId}/reply`;

        const body: Record<string, unknown> = opts.html
          ? { message: { body: { contentType: "HTML", content: opts.body } } }
          : { comment: opts.body };

        await graph.post<void>(endpoint, body);
        outputData({ message: "Reply sent." });
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 mail delete <messageId>
  mail
    .command("delete <messageId>")
    .description("Delete a message")
    .action(async (messageId) => {
      try {
        validateId(messageId, "message ID");
        await graph.delete(`/me/messages/${messageId}`);
        outputData({ message: `Message ${messageId} deleted.` });
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 mail delta
  // NOTE: Output defaults to NDJSON (one JSON object per line) rather than the
  // standard { "data": [...] } envelope used by other commands. This is
  // intentional: watcher processes pipe this to `jq -c` / while-read loops
  // where a growing JSON array is not streamable. Use --format json to get the
  // standard envelope instead.
  //
  // TODO: a future `m365 mail subscribe` command will use Graph change
  // notifications (webhooks) instead of polling. That belongs in a separate
  // command and is out of scope here.
  mail
    .command("delta")
    .description(
      "Track mail changes since last call using Graph delta query. " +
        "Persists an opaque delta link between invocations. " +
        "Designed for repeated calls by an external watcher process. " +
        "NDJSON output by default (see --format)."
    )
    .option("--folder <name|id>", "Folder name or ID", "inbox")
    .option(
      "--state-file <path>",
      "File where the delta link is persisted between runs " +
        "(default: $XDG_STATE_HOME/m365-cli/mail-delta-<folder>.link)"
    )
    .option(
      "--change-type <type>",
      "Filter by change type on initial sync: created, updated, or deleted"
    )
    .option(
      "--select <fields>",
      `Comma-separated fields to return on initial sync ` +
        `(default: ${DEFAULT_DELTA_SELECT}; internetMessageId is always appended if missing)`
    )
    .option(
      "--max-page-size <n>",
      "Max messages per page (sets Prefer: odata.maxpagesize header)",
      "50"
    )
    .option(
      "--reset",
      "Delete the state file and exit without calling Graph. Next run re-initializes."
    )
    .option(
      "--init-quiet",
      "On first run: drain all pages but suppress output, saving only the final delta link. " +
        "Prevents emitting every existing message as events on initial watcher setup."
    )
    .option(
      "--format <fmt>",
      "Output format: ndjson (default, one object per line) or json ({ data: [...] } envelope)",
      "ndjson"
    )
    .action(async (opts) => {
      try {
        const folder = resolveFolder(opts.folder ?? "inbox");
        const stateFile: string = opts.stateFile ?? getDefaultStateFile(folder);
        const isNdjson = opts.format !== "json";
        const maxPageSize = parseInt(opts.maxPageSize ?? "50", 10);

        if (opts.reset) {
          await deleteStateFile(stateFile);
          return;
        }

        const existingLink = await readStateFile(stateFile);
        const isInitialSync = existingLink === null;
        const quiet = isInitialSync && !!opts.initQuiet;
        const startUrl = existingLink ?? buildDeltaUrl(folder, opts);

        try {
          await drainPages(startUrl, stateFile, quiet, isNdjson, maxPageSize);
        } catch (err) {
          if (isDeltaExpired(err)) {
            process.stderr.write(
              JSON.stringify({
                warning: "delta token expired, resyncing",
                code: (err as GraphError).code,
              }) + "\n"
            );
            await deleteStateFile(stateFile);
            // Re-run as a quiet initial sync regardless of --init-quiet;
            // we don't want to flood the watcher with the entire inbox
            // just because a token expired mid-day.
            await drainPages(
              buildDeltaUrl(folder, opts),
              stateFile,
              true,
              isNdjson,
              maxPageSize
            );
            return;
          }
          throw err;
        }
      } catch (err) {
        handleCommandError(err);
      }
    });

  // m365 mail folders
  const folders = mail.command("folders").description("Mail folder operations");

  folders
    .command("list")
    .description("List mail folders")
    .action(async () => {
      try {
        const result = await graph.get<{ value: unknown[] }>(
          "/me/mailFolders?includeHiddenFolders=false"
        );
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  folders
    .command("get <folderId>")
    .description("Get a mail folder by ID")
    .action(async (folderId) => {
      try {
        validateId(folderId, "folder ID");
        const result = await graph.get<unknown>(`/me/mailFolders/${folderId}`);
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });
}
