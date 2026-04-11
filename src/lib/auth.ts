import {
  PublicClientApplication,
  ConfidentialClientApplication,
  DeviceCodeRequest,
  AuthenticationResult,
  Configuration,
  ICachePlugin,
  TokenCacheContext,
} from "@azure/msal-node";
import fs from "fs";
import path from "path";
import os from "os";
import { readConfig, getConfigDir, isProfileActive } from "./config.js";

// Client credentials (app-only) must use /.default — individual scopes are for delegated flows
const GRAPH_SCOPES = ["https://graph.microsoft.com/.default"];

// Lists.ReadWrite.All does not exist as a delegated permission — lists are covered by Sites.ReadWrite.All
const DELEGATED_SCOPES = [
  "https://graph.microsoft.com/Sites.ReadWrite.All",
  "https://graph.microsoft.com/Files.ReadWrite.All",
  "https://graph.microsoft.com/Mail.ReadWrite",
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/Calendars.ReadWrite",
  "https://graph.microsoft.com/Calendars.ReadWrite.Shared",
  "https://graph.microsoft.com/Team.ReadBasic.All",
  "https://graph.microsoft.com/Channel.ReadBasic.All",
  "https://graph.microsoft.com/ChannelMessage.Send",
  "https://graph.microsoft.com/ChannelMessage.Read.All",
  "https://graph.microsoft.com/Chat.ReadWrite",
  "offline_access",
];

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const DOMAIN_RE = /^[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?)+$/;

function validateTenantId(tenantId: string): void {
  if (!UUID_RE.test(tenantId) && !DOMAIN_RE.test(tenantId)) {
    throw new Error(`Invalid tenant ID "${tenantId}": must be a GUID or domain name (e.g. contoso.onmicrosoft.com).`);
  }
}

const KEYTAR_SERVICE = "m365-cli";
const LEGACY_KEYTAR_SERVICE = "sp-cli";
const KEYTAR_ACCOUNT = "access-token";

// Legacy paths kept as constants for backward-compat / migration fallback.
const LEGACY_CLI_DIR = path.join(os.homedir(), ".sp-cli");
const LEGACY_TOKEN_FILE = path.join(LEGACY_CLI_DIR, "token");
const LEGACY_MSAL_CACHE_FILE = path.join(LEGACY_CLI_DIR, "msal-cache.json");

// Dynamic helpers — respect the active profile directory.
function getMsalCacheFile(): string {
  return path.join(getConfigDir(), "msal-cache.json");
}

function getActiveTokenFile(): string {
  return path.join(getConfigDir(), "token");
}

function createCachePlugin(): ICachePlugin {
  return {
    beforeCacheAccess: async (cacheContext: TokenCacheContext) => {
      const primaryFile = getMsalCacheFile();
      // When a profile is active, only use the profile's cache file.
      // Otherwise fall back to the legacy path for migration purposes.
      const file = fs.existsSync(primaryFile) ? primaryFile
        : (!isProfileActive() && fs.existsSync(LEGACY_MSAL_CACHE_FILE)) ? LEGACY_MSAL_CACHE_FILE
        : null;
      if (file) {
        try { cacheContext.tokenCache.deserialize(fs.readFileSync(file, "utf8")); } catch {}
      }
    },
    afterCacheAccess: async (cacheContext: TokenCacheContext) => {
      if (cacheContext.cacheHasChanged) {
        const configDir = getConfigDir();
        const cacheFile = getMsalCacheFile();
        try {
          fs.mkdirSync(configDir, { recursive: true, mode: 0o700 });
          fs.writeFileSync(cacheFile, cacheContext.tokenCache.serialize(), {
            encoding: "utf8",
            mode: 0o600,
          });
        } catch {}
      }
    },
  };
}

async function getKeytar() {
  try {
    const keytar = await import("keytar");
    // Verify keytar actually works (fails silently on WSL)
    const k = keytar.default ?? keytar;
    await k.getPassword("__m365-cli-test__", "__test__");
    return k;
  } catch {
    return null;
  }
}

function storeTokenFile(token: string): void {
  const dir = getConfigDir();
  fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
  fs.writeFileSync(getActiveTokenFile(), token, { encoding: "utf8", mode: 0o600 });
}

function getTokenFileContent(): string | null {
  const primaryFile = getActiveTokenFile();
  // For profile mode, only use the profile's token file (no legacy fallback).
  const files = isProfileActive() ? [primaryFile] : [primaryFile, LEGACY_TOKEN_FILE];
  for (const f of files) {
    try {
      if (fs.existsSync(f)) return fs.readFileSync(f, "utf8").trim();
    } catch {}
  }
  return null;
}

function deleteTokenFiles(): void {
  const primaryFile = getActiveTokenFile();
  const files = isProfileActive() ? [primaryFile] : [primaryFile, LEGACY_TOKEN_FILE];
  for (const f of files) {
    try { if (fs.existsSync(f)) fs.unlinkSync(f); } catch {}
  }
}

export async function storeToken(token: string): Promise<void> {
  // Profiles always use file-based storage to keep each profile isolated.
  if (isProfileActive()) {
    storeTokenFile(token);
    return;
  }
  const keytar = await getKeytar();
  if (keytar) {
    await keytar.setPassword(KEYTAR_SERVICE, KEYTAR_ACCOUNT, token);
  } else {
    storeTokenFile(token);
  }
}

export async function getStoredToken(): Promise<string | null> {
  // Profiles always use file-based storage.
  if (isProfileActive()) {
    return getTokenFileContent();
  }
  const keytar = await getKeytar();
  if (keytar) {
    return (await keytar.getPassword(KEYTAR_SERVICE, KEYTAR_ACCOUNT))
      ?? (await keytar.getPassword(LEGACY_KEYTAR_SERVICE, KEYTAR_ACCOUNT));
  }
  return getTokenFileContent();
}

export async function deleteStoredToken(): Promise<void> {
  // Profiles always use file-based storage.
  if (isProfileActive()) {
    deleteTokenFiles();
    return;
  }
  const keytar = await getKeytar();
  if (keytar) {
    await keytar.deletePassword(KEYTAR_SERVICE, KEYTAR_ACCOUNT);
    await keytar.deletePassword(LEGACY_KEYTAR_SERVICE, KEYTAR_ACCOUNT);
  }
  deleteTokenFiles();
}

export async function acquireTokenSilent(
  tenantId: string,
  clientId: string
): Promise<string | null> {
  try {
    validateTenantId(tenantId);
    const msalConfig: Configuration = {
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
      cache: { cachePlugin: createCachePlugin() },
    };
    const pca = new PublicClientApplication(msalConfig);
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts.length === 0) return null;
    const result = await pca.acquireTokenSilent({
      scopes: DELEGATED_SCOPES,
      account: accounts[0],
    });
    return result?.accessToken ?? null;
  } catch {
    return null;
  }
}

export async function getAccessToken(): Promise<string> {
  // 1. Env var override
  if (process.env.SP_CLI_ACCESS_TOKEN) {
    return process.env.SP_CLI_ACCESS_TOKEN;
  }

  const config = readConfig();

  // 2. Client credentials (service principal)
  if (config.tenantId && config.clientId && config.clientSecret) {
    return getClientCredentialToken(
      config.tenantId,
      config.clientId,
      config.clientSecret
    );
  }

  // 3. MSAL silent refresh (uses persisted refresh token from prior device code login)
  if (config.tenantId && config.clientId) {
    const silent = await acquireTokenSilent(config.tenantId, config.clientId);
    if (silent) return silent;
  }

  // 4. Stored raw token (legacy fallback)
  const stored = await getStoredToken();
  if (stored) return stored;

  throw new Error(
    "Not authenticated. Run `m365 auth setup` then `m365 auth login`, or set SP_CLI_ACCESS_TOKEN."
  );
}

export async function getClientCredentialToken(
  tenantId: string,
  clientId: string,
  clientSecret: string
): Promise<string> {
  validateTenantId(tenantId);
  const msalConfig: Configuration = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      clientSecret,
    },
  };

  const cca = new ConfidentialClientApplication(msalConfig);
  const result: AuthenticationResult | null = await cca.acquireTokenByClientCredential({
    scopes: GRAPH_SCOPES,
  });

  if (!result?.accessToken) {
    throw new Error("Failed to acquire token via client credentials.");
  }
  return result.accessToken;
}

export async function deviceCodeLogin(
  tenantId: string,
  clientId: string
): Promise<string> {
  validateTenantId(tenantId);
  const msalConfig: Configuration = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
  };

  const pca = new PublicClientApplication({
    ...msalConfig,
    cache: { cachePlugin: createCachePlugin() },
  });

  const deviceCodeRequest: DeviceCodeRequest = {
    scopes: DELEGATED_SCOPES,
    deviceCodeCallback: (response) => {
      console.error(response.message);
    },
  };

  const result = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
  if (!result?.accessToken) {
    throw new Error("Device code login failed: no access token returned.");
  }
  return result.accessToken;
}
