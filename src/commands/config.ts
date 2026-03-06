import { Command } from "commander";
import { readConfig, mergeConfig } from "../lib/config.js";
import { outputData, handleCommandError } from "../lib/output.js";

export function registerConfigCommands(program: Command): void {
  const config = program.command("config").description("Manage CLI configuration");

  config
    .command("get")
    .description("Display current configuration")
    .action(() => {
      try {
        const cfg = readConfig();
        // Mask client secret
        const display = {
          ...cfg,
          clientSecret: cfg.clientSecret ? "***" : undefined,
        };
        outputData(display);
      } catch (err) {
        handleCommandError(err);
      }
    });

  config
    .command("set")
    .description("Update configuration values")
    .option("--tenant <url>", "Default SharePoint tenant URL")
    .option("--site <siteId>", "Default site ID")
    .option("--drive <driveId>", "Default drive ID")
    .option("--tenant-id <tenantId>", "Azure AD tenant ID")
    .option("--client-id <clientId>", "App client ID")
    .action((opts) => {
      try {
        const updates: Record<string, string> = {};
        if (opts.tenant) updates.tenantUrl = opts.tenant;
        if (opts.site) updates.defaultSiteId = opts.site;
        if (opts.drive) updates.defaultDriveId = opts.drive;
        if (opts.tenantId) updates.tenantId = opts.tenantId;
        if (opts.clientId) updates.clientId = opts.clientId;

        if (Object.keys(updates).length === 0) {
          throw new Error("No configuration values provided. Use --help for options.");
        }

        const merged = mergeConfig(updates);
        outputData({ message: "Configuration updated.", config: { ...merged, clientSecret: merged.clientSecret ? "***" : undefined } });
      } catch (err) {
        handleCommandError(err);
      }
    });
}
