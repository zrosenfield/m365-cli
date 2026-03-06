import { Command } from "commander";
import { graph, validateId } from "../lib/graph.js";
import { outputData, handleCommandError } from "../lib/output.js";

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
