import { Command } from "commander";
import { graph } from "../lib/graph.js";
import { readConfig } from "../lib/config.js";
import { outputData, handleCommandError } from "../lib/output.js";

function resolveSite(opts: { site?: string }): string {
  const siteId = opts.site || readConfig().defaultSiteId;
  if (!siteId) throw new Error("Site ID required. Use --site or run `sp config set --site <id>`.");
  return siteId;
}

export function registerDriveCommands(program: Command): void {
  const drives = program.command("drives").description("SharePoint drive (document library) operations");

  drives
    .command("list")
    .description("List drives in a site")
    .option("--site <id>", "Site ID")
    .action(async (opts) => {
      try {
        const siteId = resolveSite(opts);
        const result = await graph.get<{ value: unknown[] }>(`/sites/${siteId}/drives`);
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  drives
    .command("get <driveId>")
    .description("Get a specific drive by ID")
    .option("--site <id>", "Site ID")
    .action(async (driveId, opts) => {
      try {
        const siteId = resolveSite(opts);
        const result = await graph.get<unknown>(`/sites/${siteId}/drives/${driveId}`);
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });
}
