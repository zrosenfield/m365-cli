#!/usr/bin/env node
import { Command } from "commander";
import { registerAuthCommands } from "./commands/auth.js";
import { registerConfigCommands } from "./commands/config.js";
import { registerSiteCommands } from "./commands/sites.js";
import { registerDriveCommands } from "./commands/drives.js";
import { registerFileCommands } from "./commands/files.js";
import { registerListCommands } from "./commands/lists.js";
import { registerPermissionCommands } from "./commands/permissions.js";
import { registerMailCommands } from "./commands/mail.js";
import { registerCalendarCommands } from "./commands/calendar.js";

const program = new Command();

program
  .name("m365")
  .description("Microsoft 365 CLI for AI agents")
  .version("0.1.0");

registerAuthCommands(program);
registerConfigCommands(program);
registerSiteCommands(program);
registerDriveCommands(program);
registerFileCommands(program);
registerListCommands(program);
registerPermissionCommands(program);
registerMailCommands(program);
registerCalendarCommands(program);

program.parseAsync(process.argv).catch(() => {
  // Errors are handled and printed inside each command via handleCommandError.
  // This catch exists only to silence Commander's unhandled-rejection warning.
  process.exit(1);
});
