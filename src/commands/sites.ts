import { Command } from "commander";
import { graph, validateId } from "../lib/graph.js";
import { readConfig } from "../lib/config.js";
import { outputData, handleCommandError } from "../lib/output.js";

function resolveTenant(opts: { tenant?: string }): string {
  const url = opts.tenant || readConfig().tenantUrl;
  if (!url) throw new Error("Tenant URL required. Use --tenant or run `sp config set --tenant <url>`.");
  return url.replace(/\/$/, "");
}

export function registerSiteCommands(program: Command): void {
  const sites = program.command("sites").description("SharePoint site operations");

  sites
    .command("list")
    .description("List all SharePoint sites in the tenant")
    .option("--tenant <url>", "SharePoint tenant URL override")
    .action(async (opts) => {
      try {
        const result = await graph.get<{ value: unknown[] }>(`/sites?search=*`);
        outputData(result.value);
      } catch (err) {
        handleCommandError(err);
      }
    });

  sites
    .command("get [siteId]")
    .description("Get a specific site by ID or URL")
    .option("--tenant <url>", "SharePoint tenant URL override")
    .option("--url <siteUrl>", "Full SharePoint site URL to look up")
    .action(async (siteId, opts) => {
      try {
        let path: string;
        if (opts.url) {
          const parsed = new URL(opts.url);
          const hostname = parsed.hostname;
          const sitePath = parsed.pathname;
          path = `/sites/${hostname}:${sitePath}`;
        } else if (siteId) {
          validateId(siteId, "site ID");
          path = `/sites/${siteId}`;
        } else {
          const tenantUrl = resolveTenant(opts);
          const parsed = new URL(tenantUrl);
          path = `/sites/${parsed.hostname}:/`;
        }
        const result = await graph.get<unknown>(path);
        outputData(result);
      } catch (err) {
        handleCommandError(err);
      }
    });
}
