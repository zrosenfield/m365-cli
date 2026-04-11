import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { readConfig, setProfileDir, getConfigDir, isProfileActive } from "../../src/lib/config.js";

// Env var keys managed by readConfig
const ENV_KEYS = [
  "SP_CLI_TENANT_ID",
  "SP_CLI_CLIENT_ID",
  "SP_CLI_CLIENT_SECRET",
  "SP_CLI_TENANT_URL",
  "SP_CLI_SITE_ID",
  "SP_CLI_DRIVE_ID",
] as const;

function clearEnv() {
  for (const k of ENV_KEYS) delete process.env[k];
}

describe("readConfig env var overrides", () => {
  beforeEach(clearEnv);
  afterEach(clearEnv);

  it("does not throw when no env vars are set", () => {
    // readConfig reads ~/.m365-cli/config.json (or ~/.sp-cli for legacy) which may or may not exist.
    // We only verify it returns without throwing.
    expect(() => readConfig()).not.toThrow();
  });

  it("picks up SP_CLI_TENANT_ID", () => {
    process.env.SP_CLI_TENANT_ID = "test-tenant";
    expect(readConfig().tenantId).toBe("test-tenant");
  });

  it("picks up SP_CLI_CLIENT_ID", () => {
    process.env.SP_CLI_CLIENT_ID = "test-client";
    expect(readConfig().clientId).toBe("test-client");
  });

  it("picks up SP_CLI_SITE_ID and SP_CLI_DRIVE_ID", () => {
    process.env.SP_CLI_SITE_ID = "site-abc";
    process.env.SP_CLI_DRIVE_ID = "drive-xyz";
    const cfg = readConfig();
    expect(cfg.defaultSiteId).toBe("site-abc");
    expect(cfg.defaultDriveId).toBe("drive-xyz");
  });

  it("picks up SP_CLI_TENANT_URL", () => {
    process.env.SP_CLI_TENANT_URL = "https://contoso.sharepoint.com";
    expect(readConfig().tenantUrl).toBe("https://contoso.sharepoint.com");
  });
});

describe("profile dir (setProfileDir / getConfigDir / isProfileActive)", () => {
  afterEach(() => {
    // Always reset to default (null) so profile state does not leak between tests.
    setProfileDir(null);
  });

  it("isProfileActive returns false when no profile is set", () => {
    expect(isProfileActive()).toBe(false);
  });

  it("isProfileActive returns true after setProfileDir", () => {
    setProfileDir("/tmp/test-profile");
    expect(isProfileActive()).toBe(true);
  });

  it("getConfigDir returns the profile dir when one is set", () => {
    setProfileDir("/tmp/my-profile");
    expect(getConfigDir()).toBe("/tmp/my-profile");
  });

  it("getConfigDir returns a non-profile path after profile is cleared", () => {
    setProfileDir("/tmp/my-profile");
    setProfileDir(null);
    // After clearing, it should NOT return the old profile dir.
    expect(getConfigDir()).not.toBe("/tmp/my-profile");
    expect(isProfileActive()).toBe(false);
  });
});
