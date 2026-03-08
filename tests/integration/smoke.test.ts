/**
 * Integration smoke tests — real Microsoft 365 tenant.
 *
 * These tests are SKIPPED automatically when credentials are absent (forks,
 * local dev without a dev tenant, etc.).
 *
 * Required env vars (GitHub Actions secrets):
 *   SP_CLI_TENANT_ID, SP_CLI_CLIENT_ID, SP_CLI_CLIENT_SECRET
 *
 * Optional (needed for files list):
 *   SP_CLI_SITE_ID, SP_CLI_DRIVE_ID
 *
 * NOTE: Mail and calendar commands use /me/ endpoints which require delegated
 * auth (device code). They are NOT tested here because CI uses a service
 * principal (app-only) token that cannot access /me/ endpoints.
 *
 * Run manually:
 *   SP_CLI_TENANT_ID=... SP_CLI_CLIENT_ID=... SP_CLI_CLIENT_SECRET=... \
 *     npm run test:integration
 */

import { describe, it, expect, beforeAll, vi, afterEach } from "vitest";
import { Command } from "commander";
import { registerSiteCommands } from "../../src/commands/sites.js";
import { registerDriveCommands } from "../../src/commands/drives.js";
import { registerFileCommands } from "../../src/commands/files.js";

const hasCredentials = !!(
  process.env.SP_CLI_TENANT_ID &&
  process.env.SP_CLI_CLIENT_ID &&
  process.env.SP_CLI_CLIENT_SECRET
);

// --- helpers ---

async function runCommand(
  register: (p: Command) => void,
  args: string[]
): Promise<{ stdout: string; stderr: string }> {
  let stdout = "";
  let stderr = "";

  const stdoutSpy = vi
    .spyOn(process.stdout, "write")
    .mockImplementation((data) => {
      stdout += String(data);
      return true;
    });
  const stderrSpy = vi
    .spyOn(process.stderr, "write")
    .mockImplementation((data) => {
      stderr += String(data);
      return true;
    });

  const program = new Command();
  program.exitOverride();
  register(program);

  try {
    await program.parseAsync(["node", "m365", ...args]);
  } finally {
    stdoutSpy.mockRestore();
    stderrSpy.mockRestore();
  }

  return { stdout, stderr };
}

// --- tests ---

describe.skipIf(!hasCredentials)("integration smoke tests", () => {
  let firstSiteId: string;
  let firstDriveId: string;

  beforeAll(async () => {
    // Discover a site and drive for subsequent tests
    const sitesResult = await runCommand(registerSiteCommands, ["sites", "list"]);
    const sites = JSON.parse(sitesResult.stdout).data as { id: string }[];
    expect(sites.length).toBeGreaterThan(0);
    firstSiteId = sites[0].id;

    const drivesResult = await runCommand(
      (p) => registerDriveCommands(p),
      ["drives", "list", "--site", firstSiteId]
    );
    const drives = JSON.parse(drivesResult.stdout).data as { id: string }[];
    expect(drives.length).toBeGreaterThan(0);
    firstDriveId = drives[0].id;
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("sites list returns { data: [...] } with at least one site", async () => {
    const { stdout } = await runCommand(registerSiteCommands, ["sites", "list"]);
    const parsed = JSON.parse(stdout);
    expect(parsed).toHaveProperty("data");
    expect(Array.isArray(parsed.data)).toBe(true);
    expect(parsed.data.length).toBeGreaterThan(0);
  });

  it("drives list returns { data: [...] } for the first site", async () => {
    const { stdout } = await runCommand(
      registerDriveCommands,
      ["drives", "list", "--site", firstSiteId]
    );
    const parsed = JSON.parse(stdout);
    expect(parsed).toHaveProperty("data");
    expect(Array.isArray(parsed.data)).toBe(true);
    expect(parsed.data.length).toBeGreaterThan(0);
  });

  it("files list returns { data: [...] } for the first drive", async () => {
    const { stdout } = await runCommand(
      registerFileCommands,
      ["files", "list", "--site", firstSiteId, "--drive", firstDriveId]
    );
    const parsed = JSON.parse(stdout);
    expect(parsed).toHaveProperty("data");
    expect(Array.isArray(parsed.data)).toBe(true);
  });
});
