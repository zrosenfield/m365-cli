#!/usr/bin/env node
import updateNotifier from "update-notifier";
import pkg from "../package.json";
updateNotifier({ pkg }).notify();

import path from "path";
import os from "os";
import { Command } from "commander";
import { setProfileDir } from "./lib/config.js";
import { registerAuthCommands } from "./commands/auth.js";
import { registerConfigCommands } from "./commands/config.js";
import { registerSiteCommands } from "./commands/sites.js";
import { registerDriveCommands } from "./commands/drives.js";
import { registerFileCommands } from "./commands/files.js";
import { registerListCommands } from "./commands/lists.js";
import { registerPermissionCommands } from "./commands/permissions.js";
import { registerMailCommands } from "./commands/mail.js";
import { registerCalendarCommands } from "./commands/calendar.js";
import { registerTeamsCommands } from "./commands/teams.js";

// Extract --profile <name> from argv before Commander sees it so it works at
// any position in the command line (e.g. `m365 files list --profile patty`).
const rawArgv = process.argv.slice(2);
const profileFlagIdx = rawArgv.findIndex((a) => a === "--profile");
if (profileFlagIdx !== -1 && profileFlagIdx + 1 < rawArgv.length) {
  const profileName = rawArgv[profileFlagIdx + 1];
  rawArgv.splice(profileFlagIdx, 2);
  const profileDir = path.join(os.homedir(), ".m365-cli", "profiles", profileName);
  setProfileDir(profileDir);
}

const program = new Command();

program
  .name("m365")
  .description("Microsoft 365 CLI for AI agents")
  .version("0.1.0")
  // Documented here for --help; actual extraction is done above before parsing.
  .option("--profile <name>", "Use a named profile (~/.m365-cli/profiles/<name>/)");

registerAuthCommands(program);
registerConfigCommands(program);
registerSiteCommands(program);
registerDriveCommands(program);
registerFileCommands(program);
registerListCommands(program);
registerPermissionCommands(program);
registerMailCommands(program);
registerCalendarCommands(program);
registerTeamsCommands(program);

program.parseAsync(["node", "m365", ...rawArgv]).catch(() => {
  // Errors are handled and printed inside each command via handleCommandError.
  // This catch exists only to silence Commander's unhandled-rejection warning.
  process.exit(1);
});
