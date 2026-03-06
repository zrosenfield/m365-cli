import fs from "fs";
import path from "path";
import os from "os";

const CONFIG_DIR = path.join(os.homedir(), ".sp-cli");
const CONFIG_FILE = path.join(CONFIG_DIR, "config.json");

export interface SpConfig {
  tenantId?: string;
  clientId?: string;
  clientSecret?: string;
  tenantUrl?: string;
  defaultSiteId?: string;
  defaultDriveId?: string;
}

export function readConfig(): SpConfig {
  let file: SpConfig = {};
  try {
    if (fs.existsSync(CONFIG_FILE)) {
      file = JSON.parse(fs.readFileSync(CONFIG_FILE, "utf8")) as SpConfig;
    }
  } catch {
    // ignore parse errors
  }

  // Env vars override file config
  return {
    ...file,
    ...(process.env.SP_CLI_TENANT_ID && { tenantId: process.env.SP_CLI_TENANT_ID }),
    ...(process.env.SP_CLI_CLIENT_ID && { clientId: process.env.SP_CLI_CLIENT_ID }),
    ...(process.env.SP_CLI_CLIENT_SECRET && { clientSecret: process.env.SP_CLI_CLIENT_SECRET }),
    ...(process.env.SP_CLI_TENANT_URL && { tenantUrl: process.env.SP_CLI_TENANT_URL }),
    ...(process.env.SP_CLI_SITE_ID && { defaultSiteId: process.env.SP_CLI_SITE_ID }),
    ...(process.env.SP_CLI_DRIVE_ID && { defaultDriveId: process.env.SP_CLI_DRIVE_ID }),
  };
}

export function writeConfig(config: SpConfig): void {
  if (!fs.existsSync(CONFIG_DIR)) {
    fs.mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
  }
  fs.writeFileSync(CONFIG_FILE, JSON.stringify(config, null, 2), {
    encoding: "utf8",
    mode: 0o600,
  });
}

export function mergeConfig(updates: Partial<SpConfig>): SpConfig {
  const existing = readConfig();
  const merged = { ...existing, ...updates };
  writeConfig(merged);
  return merged;
}

export function clearConfig(): void {
  if (fs.existsSync(CONFIG_FILE)) {
    fs.unlinkSync(CONFIG_FILE);
  }
}
