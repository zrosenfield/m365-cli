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
import { readConfig } from "./config.js";

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

const KEYTAR_SERVICE = "sp-cli";
const KEYTAR_ACCOUNT = "access-token";
const TOKEN_FILE = path.join(os.homedir(), ".sp-cli", "token");
const MSAL_CACHE_FILE = path.join(os.homedir(), ".sp-cli", "msal-cache.json");

function createCachePlugin(): ICachePlugin {
  return {
    beforeCacheAccess: async (cacheContext: TokenCacheContext) => {
      try {
        if (fs.existsSync(MSAL_CACHE_FILE)) {
          const data = fs.readFileSync(MSAL_CACHE_FILE, "utf8");
          cacheContext.tokenCache.deserialize(data);
        }
      } catch {}
    },
    afterCacheAccess: async (cacheContext: TokenCacheContext) => {
      if (cacheContext.cacheHasChanged) {
        try {
          const dir = path.dirname(MSAL_CACHE_FILE);
          fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
          fs.writeFileSync(MSAL_CACHE_FILE, cacheContext.tokenCache.serialize(), {
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
    await k.getPassword("__sp-cli-test__", "__test__");
    return k;
  } catch {
    return null;
  }
}

function storeTokenFile(token: string): void {
  const dir = path.dirname(TOKEN_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
  fs.writeFileSync(TOKEN_FILE, token, { encoding: "utf8", mode: 0o600 });
}

function getTokenFile(): string | null {
  try {
    if (fs.existsSync(TOKEN_FILE)) return fs.readFileSync(TOKEN_FILE, "utf8").trim();
  } catch {}
  return null;
}

function deleteTokenFile(): void {
  try {
    if (fs.existsSync(TOKEN_FILE)) fs.unlinkSync(TOKEN_FILE);
  } catch {}
}

export async function storeToken(token: string): Promise<void> {
  const keytar = await getKeytar();
  if (keytar) {
    await keytar.setPassword(KEYTAR_SERVICE, KEYTAR_ACCOUNT, token);
  } else {
    storeTokenFile(token);
  }
}

export async function getStoredToken(): Promise<string | null> {
  const keytar = await getKeytar();
  if (keytar) {
    return keytar.getPassword(KEYTAR_SERVICE, KEYTAR_ACCOUNT);
  }
  return getTokenFile();
}

export async function deleteStoredToken(): Promise<void> {
  const keytar = await getKeytar();
  if (keytar) {
    await keytar.deletePassword(KEYTAR_SERVICE, KEYTAR_ACCOUNT);
  }
  deleteTokenFile();
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
    "Not authenticated. Run `sp auth setup` then `sp auth login`, or set SP_CLI_ACCESS_TOKEN."
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
