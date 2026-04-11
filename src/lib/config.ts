import fs from "fs";
import path from "path";
import os from "os";

const NEW_CONFIG_DIR = path.join(os.homedir(), ".m365-cli");
const LEGACY_CONFIG_DIR = path.join(os.homedir(), ".sp-cli");

// When set, all config reads/writes use this directory instead of the defaults.
let _profileDir: string | null = null;

/** Activate a named profile directory.  Pass null to deactivate. */
export function setProfileDir(dir: string | null): void {
  _profileDir = dir;
}

/** Returns true when a named profile is currently active. */
export function isProfileActive(): boolean {
  return _profileDir !== null;
}

function resolveDefaultConfigDir(): string {
  // Check for config.json specifically — the directory may exist (e.g. from
  // creating a profiles subdirectory) without a config file inside it.
  if (fs.existsSync(path.join(NEW_CONFIG_DIR, "config.json"))) return NEW_CONFIG_DIR;
  if (fs.existsSync(path.join(LEGACY_CONFIG_DIR, "config.json"))) return LEGACY_CONFIG_DIR;
  return NEW_CONFIG_DIR;
}

/**
 * Returns the active configuration directory.
 *
 * Priority: profile dir (if set) → ~/.m365-cli (if exists) → ~/.sp-cli
 * (migration fallback) → ~/.m365-cli (default for new installs).
 */
export function getConfigDir(): string {
  if (_profileDir !== null) return _profileDir;
  return resolveDefaultConfigDir();
}

export interface SpConfig {
  tenantId?: string;
  clientId?: string;
  clientSecret?: string;
  tenantUrl?: string;
  defaultSiteId?: string;
  defaultDriveId?: string;
}

export function readConfig(): SpConfig {
  const configFile = path.join(getConfigDir(), "config.json");
  let file: SpConfig = {};
  try {
    if (fs.existsSync(configFile)) {
      file = JSON.parse(fs.readFileSync(configFile, "utf8")) as SpConfig;
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
  const configDir = getConfigDir();
  fs.mkdirSync(configDir, { recursive: true, mode: 0o700 });
  fs.writeFileSync(path.join(configDir, "config.json"), JSON.stringify(config, null, 2), {
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
  if (_profileDir !== null) {
    const f = path.join(_profileDir, "config.json");
    if (fs.existsSync(f)) fs.unlinkSync(f);
    return;
  }
  for (const dir of [NEW_CONFIG_DIR, LEGACY_CONFIG_DIR]) {
    const f = path.join(dir, "config.json");
    if (fs.existsSync(f)) fs.unlinkSync(f);
  }
}
