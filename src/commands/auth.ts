import { Command } from "commander";
import * as readline from "readline";
import { readConfig, mergeConfig } from "../lib/config.js";
import {
  deviceCodeLogin,
  storeToken,
  deleteStoredToken,
  getAccessToken,
} from "../lib/auth.js";
import { outputData, outputError, handleCommandError } from "../lib/output.js";

function prompt(question: string): Promise<string> {
  const rl = readline.createInterface({ input: process.stdin, output: process.stderr });
  return new Promise((resolve) => {
    rl.question(question, (answer) => {
      rl.close();
      resolve(answer.trim());
    });
  });
}

export function registerAuthCommands(program: Command): void {
  const auth = program.command("auth").description("Authentication commands");

  auth
    .command("setup")
    .description("Configure tenant ID, client ID, and optionally client secret")
    .option("--tenant-id <tenantId>", "Azure AD tenant ID (GUID or domain)")
    .option("--client-id <clientId>", "App registration client ID")
    .option("--client-secret <secret>", "Client secret (for service principal auth)")
    .option("--tenant-url <url>", "SharePoint tenant URL (e.g. https://contoso.sharepoint.com)")
    .action(async (opts) => {
      try {
        const tenantId = opts.tenantId || (await prompt("Tenant ID: "));
        const clientId = opts.clientId || (await prompt("Client ID: "));
        const clientSecret = opts.clientSecret || (await prompt("Client secret (leave blank for device code): ")) || undefined;
        const tenantUrl = opts.tenantUrl || (await prompt("SharePoint tenant URL (optional): ")) || undefined;

        mergeConfig({
          tenantId,
          clientId,
          ...(clientSecret ? { clientSecret } : {}),
          ...(tenantUrl ? { tenantUrl } : {}),
        });

        outputData({ message: "Configuration saved." });
      } catch (err) {
        handleCommandError(err);
      }
    });

  auth
    .command("login")
    .description("Authenticate via device code interactive flow")
    .action(async () => {
      try {
        const config = readConfig();
        if (!config.tenantId || !config.clientId) {
          throw new Error("Run `sp auth setup` first to configure tenantId and clientId.");
        }
        const token = await deviceCodeLogin(config.tenantId, config.clientId);
        await storeToken(token);
        outputData({ message: "Login successful. Token stored in OS keychain." });
      } catch (err) {
        handleCommandError(err);
      }
    });

  auth
    .command("logout")
    .description("Clear stored authentication tokens")
    .action(async () => {
      try {
        await deleteStoredToken();
        outputData({ message: "Logged out. Stored token cleared." });
      } catch (err) {
        handleCommandError(err);
      }
    });

  auth
    .command("token")
    .description("Print the current access token")
    .option("--raw", "Output bare token without JSON wrapper")
    .action(async (opts) => {
      try {
        const token = await getAccessToken();
        if (opts.raw) {
          process.stdout.write(token + "\n");
        } else {
          outputData({ token });
        }
      } catch (err) {
        handleCommandError(err);
      }
    });
}
