import { Command } from "commander";
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

export function registerPermissionCommands(program: Command): void {
  const permissions = program.command("permissions").description("File permission operations");

  permissions
    .command("list <itemId>")
    .description("List permissions on a file or folder")
    .option("--site <id>", "Site ID")
    .option("--drive <id>", "Drive ID")
    .action(async (itemId, opts) => {
      try {
        const { driveId } = resolveDrive(opts);
        validateId(itemId, "item ID");
        const result = await graph.get<{ value: unknown[] }>(
          `/drives/${driveId}/items/${itemId}/permissions`
        );
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  permissions
    .command("get <itemId> <permId>")
    .description("Get a specific permission")
    .option("--site <id>", "Site ID")
    .option("--drive <id>", "Drive ID")
    .action(async (itemId, permId, opts) => {
      try {
        const { driveId } = resolveDrive(opts);
        validateId(itemId, "item ID");
        validateId(permId, "permission ID");
        const result = await graph.get<unknown>(
          `/drives/${driveId}/items/${itemId}/permissions/${permId}`
        );
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  permissions
    .command("grant <itemId>")
    .description("Grant permission to users")
    .requiredOption("--emails <emails>", "Comma-separated email addresses")
    .requiredOption("--role <role>", "Role: read, write, or owner")
    .option("--site <id>", "Site ID")
    .option("--drive <id>", "Drive ID")
    .action(async (itemId, opts) => {
      try {
        const { driveId } = resolveDrive(opts);
        validateId(itemId, "item ID");
        const emails = opts.emails.split(",").map((e: string) => e.trim());

        // Map friendly role names to Graph roles
        const roleMap: Record<string, string> = {
          reader: "read",
          writer: "write",
          owner: "owner",
          read: "read",
          write: "write",
        };
        const role = roleMap[opts.role];
        if (!role) throw new Error(`Invalid role: ${opts.role}. Use reader, writer, or owner.`);

        const result = await graph.post<unknown>(
          `/drives/${driveId}/items/${itemId}/invite`,
          {
            requireSignIn: true,
            sendInvitation: false,
            roles: [role],
            recipients: emails.map((email: string) => ({ email })),
          }
        );
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  permissions
    .command("update <itemId> <permId>")
    .description("Update a permission's role")
    .requiredOption("--role <role>", "New role: read, write, or owner")
    .option("--site <id>", "Site ID")
    .option("--drive <id>", "Drive ID")
    .action(async (itemId, permId, opts) => {
      try {
        const { driveId } = resolveDrive(opts);
        validateId(itemId, "item ID");
        validateId(permId, "permission ID");
        const roleMap: Record<string, string> = {
          reader: "read",
          writer: "write",
          owner: "owner",
          read: "read",
          write: "write",
        };
        const role = roleMap[opts.role];
        if (!role) throw new Error(`Invalid role: ${opts.role}. Use reader, writer, or owner.`);
        const result = await graph.patch<unknown>(
          `/drives/${driveId}/items/${itemId}/permissions/${permId}`,
          { roles: [role] }
        );
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });

  permissions
    .command("revoke <itemId> <permId>")
    .description("Revoke a permission")
    .option("--site <id>", "Site ID")
    .option("--drive <id>", "Drive ID")
    .action(async (itemId, permId, opts) => {
      try {
        const { driveId } = resolveDrive(opts);
        validateId(itemId, "item ID");
        validateId(permId, "permission ID");
        await graph.delete(`/drives/${driveId}/items/${itemId}/permissions/${permId}`);
        outputData({ message: `Permission ${permId} revoked from item ${itemId}.` });
      } catch (err) {
        handleCommandError(err);
      }
    });
}
