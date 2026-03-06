import { Command } from "commander";
import fs from "fs";
import path from "path";
import { graph, validateId } from "../lib/graph.js";
import { readConfig } from "../lib/config.js";
import { outputData, handleCommandError } from "../lib/output.js";

function resolveDrive(opts: { site?: string; drive?: string }): { siteId: string; driveId: string } {
  const config = readConfig();
  const siteId = opts.site || config.defaultSiteId;
  const driveId = opts.drive || config.defaultDriveId;
  if (!siteId) throw new Error("Site ID required. Use --site or run `sp config set --site <id>`.");
  if (!driveId) throw new Error("Drive ID required. Use --drive or run `sp config set --drive <id>`.");
  validateId(siteId, "site ID");
  validateId(driveId, "drive ID");
  return { siteId, driveId };
}

export function registerFileCommands(program: Command): void {
  const files = program.command("files").description("SharePoint file operations");

  files
    .command("list")
    .description("List files and folders in a drive path")
    .option("--site <id>", "Site ID")
    .option("--drive <id>", "Drive ID")
    .option("--path <folderPath>", "Remote folder path (default: root)")
    .action(async (opts) => {
      try {
        const { driveId } = resolveDrive(opts);
        const folderPath = opts.path ? `:${opts.path}:` : "";
        const result = await graph.get<{ value: unknown[] }>(
          `/drives/${driveId}/root${folderPath}/children`
        );
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  files
    .command("get <itemId>")
    .description("Get metadata for a file or folder")
    .option("--site <id>", "Site ID")
    .option("--drive <id>", "Drive ID")
    .action(async (itemId, opts) => {
      try {
        const { driveId } = resolveDrive(opts);
        validateId(itemId, "item ID");
        const result = await graph.get<unknown>(`/drives/${driveId}/items/${itemId}`);
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  files
    .command("upload <localPath>")
    .description("Upload a local file to SharePoint")
    .option("--site <id>", "Site ID")
    .option("--drive <id>", "Drive ID")
    .option("--remote-path <path>", "Destination path in drive (e.g. /Documents/file.txt)")
    .action(async (localPath, opts) => {
      try {
        const { driveId } = resolveDrive(opts);
        const fileName = path.basename(localPath);
        const remotePath = opts.remotePath || `/${fileName}`;
        const data = fs.readFileSync(localPath);
        const result = await graph.upload<unknown>(
          `/drives/${driveId}/root:${remotePath}:/content`,
          data
        );
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  files
    .command("download <itemId>")
    .description("Download a file from SharePoint")
    .option("--site <id>", "Site ID")
    .option("--drive <id>", "Drive ID")
    .option("--output <localPath>", "Local path to save the file")
    .action(async (itemId, opts) => {
      try {
        const { driveId } = resolveDrive(opts);
        validateId(itemId, "item ID");

        // Get download URL
        const meta = await graph.get<{ name?: string; "@microsoft.graph.downloadUrl"?: string }>(
          `/drives/${driveId}/items/${itemId}`
        );
        const downloadUrl = meta["@microsoft.graph.downloadUrl"];
        if (!downloadUrl) throw new Error("No download URL available for this item.");

        const fetch = (await import("node-fetch")).default;
        const res = await fetch(downloadUrl);
        if (!res.ok) throw new Error(`Download failed: HTTP ${res.status}`);

        const outputPath = opts.output || path.basename(meta.name || itemId);
        const buffer = await res.buffer();
        fs.writeFileSync(outputPath, buffer);
        outputData({ message: `Downloaded to ${outputPath}`, bytes: buffer.length });
      } catch (err) {
        handleCommandError(err);
      }
    });

  files
    .command("copy <itemId>")
    .description("Copy a file to another location")
    .requiredOption("--dest-path <path>", "Destination folder path")
    .option("--dest-drive <id>", "Destination drive ID (default: same drive)")
    .option("--site <id>", "Site ID")
    .option("--drive <id>", "Drive ID")
    .option("--name <name>", "New name for the copy")
    .action(async (itemId, opts) => {
      try {
        const { driveId } = resolveDrive(opts);
        validateId(itemId, "item ID");
        const destDriveId = opts.destDrive || driveId;

        const body: Record<string, unknown> = {
          parentReference: {
            driveId: destDriveId,
            path: `/drives/${destDriveId}/root:${opts.destPath}`,
          },
        };
        if (opts.name) body.name = opts.name;

        const result = await graph.post<unknown>(
          `/drives/${driveId}/items/${itemId}/copy`,
          body
        );
        outputData(result ?? { message: "Copy initiated." });
      } catch (err) {
        handleCommandError(err);
      }
    });

  files
    .command("move <itemId>")
    .description("Move a file to another location")
    .requiredOption("--dest-path <path>", "Destination folder path")
    .option("--dest-drive <id>", "Destination drive ID (default: same drive)")
    .option("--site <id>", "Site ID")
    .option("--drive <id>", "Drive ID")
    .action(async (itemId, opts) => {
      try {
        const { driveId } = resolveDrive(opts);
        validateId(itemId, "item ID");
        const destDriveId = opts.destDrive || driveId;

        const result = await graph.patch<unknown>(`/drives/${driveId}/items/${itemId}`, {
          parentReference: {
            driveId: destDriveId,
            path: `/drives/${destDriveId}/root:${opts.destPath}`,
          },
        });
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  files
    .command("rename <itemId>")
    .description("Rename a file or folder")
    .requiredOption("--name <newName>", "New name")
    .option("--site <id>", "Site ID")
    .option("--drive <id>", "Drive ID")
    .action(async (itemId, opts) => {
      try {
        const { driveId } = resolveDrive(opts);
        validateId(itemId, "item ID");
        const result = await graph.patch<unknown>(`/drives/${driveId}/items/${itemId}`, {
          name: opts.name,
        });
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  files
    .command("delete <itemId>")
    .description("Delete a file or folder")
    .option("--site <id>", "Site ID")
    .option("--drive <id>", "Drive ID")
    .action(async (itemId, opts) => {
      try {
        const { driveId } = resolveDrive(opts);
        validateId(itemId, "item ID");
        await graph.delete(`/drives/${driveId}/items/${itemId}`);
        outputData({ message: `Item ${itemId} deleted.` });
      } catch (err) {
        handleCommandError(err);
      }
    });

  files
    .command("search <query>")
    .description("Search for files in a drive")
    .option("--site <id>", "Site ID")
    .option("--drive <id>", "Drive ID")
    .action(async (query, opts) => {
      try {
        const { driveId } = resolveDrive(opts);
        const result = await graph.get<{ value: unknown[] }>(
          `/drives/${driveId}/root/search(q='${encodeURIComponent(query)}')`
        );
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });
}
