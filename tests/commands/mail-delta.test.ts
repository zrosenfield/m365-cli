import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { Command } from "commander";
import { registerMailCommands } from "../../src/commands/mail.js";

// --- module mocks (hoisted) ---

vi.mock("../../src/lib/graph.js", () => ({
  graph: {
    get: vi.fn(),
    post: vi.fn(),
    patch: vi.fn(),
    put: vi.fn(),
    delete: vi.fn(),
  },
  GraphError: class GraphError extends Error {
    status: number;
    code: string;
    constructor(status: number, code: string, message: string) {
      super(message);
      this.status = status;
      this.code = code;
      this.name = "GraphError";
    }
  },
  validateId: vi.fn(),
}));

vi.mock("node:fs/promises", () => ({
  mkdir: vi.fn(),
  readFile: vi.fn(),
  writeFile: vi.fn(),
  rename: vi.fn(),
  unlink: vi.fn(),
}));

// --- import mocked modules ---

import { graph, GraphError } from "../../src/lib/graph.js";
import { mkdir, readFile, writeFile, rename, unlink } from "node:fs/promises";

// --- helpers ---

function makeProgram(): Command {
  const p = new Command();
  p.exitOverride();
  registerMailCommands(p);
  return p;
}

function parse(p: Command, args: string[]) {
  return p.parseAsync(["node", "m365", ...args]);
}

const FAKE_DELTA_LINK = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=abc123";
const DEFAULT_STATE_FILE_PATTERN = /mail-delta-inbox\.link$/;

// --- test suite ---

describe("mail delta", () => {
  let stdoutData: string;
  let stderrData: string;

  beforeEach(() => {
    stdoutData = "";
    stderrData = "";

    vi.spyOn(process.stdout, "write").mockImplementation((data) => {
      stdoutData += String(data);
      return true;
    });
    vi.spyOn(process.stderr, "write").mockImplementation((data) => {
      stderrData += String(data);
      return true;
    });
    vi.spyOn(process, "exit").mockImplementation((_code?: number): never => {
      throw new Error(`process.exit(${_code})`);
    });

    vi.mocked(graph.get).mockReset();
    vi.mocked(mkdir).mockResolvedValue(undefined as never);
    vi.mocked(writeFile).mockResolvedValue(undefined);
    vi.mocked(rename).mockResolvedValue(undefined);
    vi.mocked(unlink).mockResolvedValue(undefined);
    vi.mocked(readFile).mockReset();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  // Helper to simulate "no state file" (initial sync)
  function noStateFile() {
    vi.mocked(readFile).mockRejectedValue(
      Object.assign(new Error("ENOENT: no such file or directory"), { code: "ENOENT" })
    );
  }

  // Helper to simulate "state file exists with a delta link"
  function withStateFile(link: string = FAKE_DELTA_LINK) {
    vi.mocked(readFile).mockResolvedValue(link as never);
  }

  // --- initial sync ---

  describe("initial sync (no state file)", () => {
    it("emits messages as NDJSON and saves the deltaLink", async () => {
      noStateFile();
      const messages = [
        { internetMessageId: "<msg1@example.com>", subject: "Hello" },
        { internetMessageId: "<msg2@example.com>", subject: "World" },
      ];
      vi.mocked(graph.get).mockResolvedValue({
        value: messages,
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), ["mail", "delta"]);

      // Each message is a separate NDJSON line
      const lines = stdoutData.trim().split("\n");
      expect(lines).toHaveLength(2);
      expect(JSON.parse(lines[0])).toEqual(messages[0]);
      expect(JSON.parse(lines[1])).toEqual(messages[1]);

      // deltaLink saved atomically
      const writeCall = vi.mocked(writeFile).mock.calls[0];
      expect(writeCall[0] as string).toMatch(/\.tmp$/);
      expect(writeCall[1]).toBe(FAKE_DELTA_LINK);
      expect(rename).toHaveBeenCalledOnce();
    });

    it("--init-quiet suppresses stdout but still saves the deltaLink", async () => {
      noStateFile();
      vi.mocked(graph.get).mockResolvedValue({
        value: [{ internetMessageId: "<msg1@example.com>" }],
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), ["mail", "delta", "--init-quiet"]);

      expect(stdoutData).toBe("");
      expect(writeFile).toHaveBeenCalledOnce();
      const writeCall = vi.mocked(writeFile).mock.calls[0];
      expect(writeCall[1]).toBe(FAKE_DELTA_LINK);
    });

    it("builds initial URL with correct path", async () => {
      noStateFile();
      vi.mocked(graph.get).mockResolvedValue({
        value: [],
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), ["mail", "delta"]);

      const url = vi.mocked(graph.get).mock.calls[0][0] as string;
      expect(url).toContain("/me/mailFolders/inbox/messages/delta");
    });

    it("uses --folder to build URL with well-known folder name", async () => {
      noStateFile();
      vi.mocked(graph.get).mockResolvedValue({
        value: [],
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), ["mail", "delta", "--folder", "sentItems"]);

      const url = vi.mocked(graph.get).mock.calls[0][0] as string;
      expect(url).toContain("/me/mailFolders/sentitems/messages/delta");
    });
  });

  // --- subsequent sync ---

  describe("subsequent sync (state file exists)", () => {
    it("calls the deltaLink directly without rebuilding the initial URL", async () => {
      withStateFile(FAKE_DELTA_LINK);
      vi.mocked(graph.get).mockResolvedValue({
        value: [{ internetMessageId: "<new@example.com>" }],
        "@odata.deltaLink": FAKE_DELTA_LINK + "_v2",
      });

      await parse(makeProgram(), ["mail", "delta"]);

      const url = vi.mocked(graph.get).mock.calls[0][0] as string;
      expect(url).toBe(FAKE_DELTA_LINK);
    });

    it("saves the new deltaLink from the response", async () => {
      const newDeltaLink = FAKE_DELTA_LINK + "_v2";
      withStateFile(FAKE_DELTA_LINK);
      vi.mocked(graph.get).mockResolvedValue({
        value: [],
        "@odata.deltaLink": newDeltaLink,
      });

      await parse(makeProgram(), ["mail", "delta"]);

      const writeCall = vi.mocked(writeFile).mock.calls[0];
      expect(writeCall[1]).toBe(newDeltaLink);
    });
  });

  // --- pagination ---

  describe("pagination", () => {
    it("follows @odata.nextLink pages before reaching @odata.deltaLink", async () => {
      noStateFile();
      const nextLink = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$skiptoken=page2";
      vi.mocked(graph.get)
        .mockResolvedValueOnce({
          value: [{ internetMessageId: "<p1@example.com>" }],
          "@odata.nextLink": nextLink,
        })
        .mockResolvedValueOnce({
          value: [{ internetMessageId: "<p2@example.com>" }],
          "@odata.deltaLink": FAKE_DELTA_LINK,
        });

      await parse(makeProgram(), ["mail", "delta"]);

      expect(graph.get).toHaveBeenCalledTimes(2);
      expect(vi.mocked(graph.get).mock.calls[1][0]).toBe(nextLink);

      const lines = stdoutData.trim().split("\n");
      expect(lines).toHaveLength(2);
    });

    it("drains three pages before deltaLink", async () => {
      noStateFile();
      const link1 = "https://graph.microsoft.com/v1.0/...?$skiptoken=p2";
      const link2 = "https://graph.microsoft.com/v1.0/...?$skiptoken=p3";
      vi.mocked(graph.get)
        .mockResolvedValueOnce({ value: [{ id: "1" }], "@odata.nextLink": link1 })
        .mockResolvedValueOnce({ value: [{ id: "2" }], "@odata.nextLink": link2 })
        .mockResolvedValueOnce({ value: [{ id: "3" }], "@odata.deltaLink": FAKE_DELTA_LINK });

      await parse(makeProgram(), ["mail", "delta"]);

      expect(graph.get).toHaveBeenCalledTimes(3);
      const lines = stdoutData.trim().split("\n");
      expect(lines).toHaveLength(3);
    });
  });

  // --- expired token ---

  describe("expired delta token", () => {
    it("writes warning to stderr, deletes state file, and quietly resyncs", async () => {
      withStateFile(FAKE_DELTA_LINK);
      const expiredError = new GraphError(410, "syncStateNotFound", "Sync state expired");
      const newDeltaLink = FAKE_DELTA_LINK + "_fresh";

      vi.mocked(graph.get)
        .mockRejectedValueOnce(expiredError)
        .mockResolvedValueOnce({
          value: [{ internetMessageId: "<existing@example.com>" }],
          "@odata.deltaLink": newDeltaLink,
        });

      await parse(makeProgram(), ["mail", "delta"]);

      // Warning on stderr (not error envelope)
      const warning = JSON.parse(stderrData.trim());
      expect(warning.warning).toContain("delta token expired");
      expect(warning.code).toBe("syncStateNotFound");

      // State file deleted
      expect(unlink).toHaveBeenCalledOnce();

      // Resync is quiet — no stdout
      expect(stdoutData).toBe("");

      // New deltaLink persisted
      const writeCall = vi.mocked(writeFile).mock.calls[0];
      expect(writeCall[1]).toBe(newDeltaLink);
    });

    it("handles syncStateInvalid the same as syncStateNotFound", async () => {
      withStateFile(FAKE_DELTA_LINK);
      const expiredError = new GraphError(400, "syncStateInvalid", "Invalid sync state");

      vi.mocked(graph.get)
        .mockRejectedValueOnce(expiredError)
        .mockResolvedValueOnce({
          value: [],
          "@odata.deltaLink": FAKE_DELTA_LINK + "_new",
        });

      await parse(makeProgram(), ["mail", "delta"]);

      expect(stderrData).toContain("syncStateInvalid");
      expect(stdoutData).toBe("");
    });

    it("does NOT suppress resync output with --init-quiet when token expires (always quiet on resync)", async () => {
      // Even if --init-quiet wasn't passed, the resync is always quiet
      withStateFile(FAKE_DELTA_LINK);
      const expiredError = new GraphError(410, "syncStateNotFound", "expired");

      vi.mocked(graph.get)
        .mockRejectedValueOnce(expiredError)
        .mockResolvedValueOnce({
          value: [{ id: "msg1" }],
          "@odata.deltaLink": FAKE_DELTA_LINK + "_new",
        });

      // No --init-quiet flag
      await parse(makeProgram(), ["mail", "delta"]);

      // Still quiet because resync is always quiet regardless of --init-quiet
      expect(stdoutData).toBe("");
    });

    it("exits non-zero on non-expiry GraphErrors", async () => {
      withStateFile(FAKE_DELTA_LINK);
      vi.mocked(graph.get).mockRejectedValue(
        new GraphError(401, "Unauthorized", "Token expired")
      );

      await expect(parse(makeProgram(), ["mail", "delta"])).rejects.toThrow("process.exit(1)");
      expect(JSON.parse(stderrData).error.status).toBe(401);
    });
  });

  // --- --reset ---

  describe("--reset", () => {
    it("deletes the state file and exits without calling Graph", async () => {
      await parse(makeProgram(), ["mail", "delta", "--reset"]);

      expect(graph.get).not.toHaveBeenCalled();
      expect(unlink).toHaveBeenCalledOnce();
      const unlinkPath = vi.mocked(unlink).mock.calls[0][0] as string;
      expect(unlinkPath).toMatch(DEFAULT_STATE_FILE_PATTERN);
    });

    it("--reset succeeds even if state file does not exist", async () => {
      vi.mocked(unlink).mockRejectedValue(
        Object.assign(new Error("ENOENT"), { code: "ENOENT" })
      );

      // Should not throw
      await expect(
        parse(makeProgram(), ["mail", "delta", "--reset"])
      ).resolves.not.toThrow();
    });

    it("--reset with --state-file deletes the specified file", async () => {
      await parse(makeProgram(), [
        "mail", "delta", "--reset", "--state-file", "/tmp/custom.link",
      ]);

      expect(vi.mocked(unlink).mock.calls[0][0]).toBe("/tmp/custom.link");
    });
  });

  // --- --select ---

  describe("--select field handling", () => {
    it("default select includes internetMessageId", async () => {
      noStateFile();
      vi.mocked(graph.get).mockResolvedValue({
        value: [],
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), ["mail", "delta"]);

      const url = vi.mocked(graph.get).mock.calls[0][0] as string;
      expect(decodeURIComponent(url)).toContain("internetMessageId");
    });

    it("appends internetMessageId when user omits it from --select", async () => {
      noStateFile();
      vi.mocked(graph.get).mockResolvedValue({
        value: [],
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), ["mail", "delta", "--select", "from,subject"]);

      const url = vi.mocked(graph.get).mock.calls[0][0] as string;
      expect(decodeURIComponent(url)).toContain("internetMessageId");
      expect(decodeURIComponent(url)).toContain("from");
      expect(decodeURIComponent(url)).toContain("subject");
    });

    it("does not duplicate internetMessageId when user includes it", async () => {
      noStateFile();
      vi.mocked(graph.get).mockResolvedValue({
        value: [],
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), [
        "mail", "delta", "--select", "internetMessageId,subject",
      ]);

      const url = vi.mocked(graph.get).mock.calls[0][0] as string;
      const decoded = decodeURIComponent(url);
      const matches = decoded.match(/internetMessageId/gi) ?? [];
      expect(matches).toHaveLength(1);
    });
  });

  // --- @removed entries ---

  describe("@removed entries", () => {
    it("emits @removed (deleted) messages as-is", async () => {
      noStateFile();
      const removed = { id: "msg1", "@removed": { reason: "deleted" } };
      vi.mocked(graph.get).mockResolvedValue({
        value: [removed],
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), ["mail", "delta"]);

      const line = JSON.parse(stdoutData.trim());
      expect(line["@removed"]).toEqual({ reason: "deleted" });
    });
  });

  // --- atomic state file write ---

  describe("atomic state file write", () => {
    it("writes to .tmp then renames to final path", async () => {
      noStateFile();
      vi.mocked(graph.get).mockResolvedValue({
        value: [],
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      const callOrder: string[] = [];
      vi.mocked(writeFile).mockImplementation(async () => {
        callOrder.push("writeFile");
      });
      vi.mocked(rename).mockImplementation(async () => {
        callOrder.push("rename");
      });

      await parse(makeProgram(), ["mail", "delta"]);

      expect(callOrder).toEqual(["writeFile", "rename"]);

      const tmpPath = vi.mocked(rename).mock.calls[0][0] as string;
      const finalPath = vi.mocked(rename).mock.calls[0][1] as string;
      expect(tmpPath).toBe(finalPath + ".tmp");
      expect(finalPath).toMatch(DEFAULT_STATE_FILE_PATTERN);
    });

    it("creates parent directories before writing", async () => {
      noStateFile();
      vi.mocked(graph.get).mockResolvedValue({
        value: [],
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), ["mail", "delta"]);

      expect(mkdir).toHaveBeenCalledWith(expect.any(String), { recursive: true });
    });
  });

  // --- --format ---

  describe("--format", () => {
    it("--format ndjson (default) outputs one line per message", async () => {
      noStateFile();
      vi.mocked(graph.get).mockResolvedValue({
        value: [{ id: "m1" }, { id: "m2" }],
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), ["mail", "delta"]);

      const lines = stdoutData.trim().split("\n");
      expect(lines).toHaveLength(2);
      lines.forEach((line) => expect(() => JSON.parse(line)).not.toThrow());
    });

    it("--format json outputs { data: [...] } envelope", async () => {
      noStateFile();
      const messages = [{ id: "m1" }, { id: "m2" }];
      vi.mocked(graph.get).mockResolvedValue({
        value: messages,
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), ["mail", "delta", "--format", "json"]);

      const out = JSON.parse(stdoutData);
      expect(out.data).toEqual(messages);
    });

    it("--format json with --init-quiet still suppresses output", async () => {
      noStateFile();
      vi.mocked(graph.get).mockResolvedValue({
        value: [{ id: "m1" }],
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), ["mail", "delta", "--format", "json", "--init-quiet"]);

      expect(stdoutData).toBe("");
    });
  });

  // --- Prefer header ---

  describe("Prefer: odata.maxpagesize header", () => {
    it("sends Prefer header with default max-page-size of 50", async () => {
      noStateFile();
      vi.mocked(graph.get).mockResolvedValue({
        value: [],
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), ["mail", "delta"]);

      const opts = vi.mocked(graph.get).mock.calls[0][1] as {
        headers?: Record<string, string>;
      };
      expect(opts?.headers?.["Prefer"]).toBe("odata.maxpagesize=50");
    });

    it("respects --max-page-size flag", async () => {
      noStateFile();
      vi.mocked(graph.get).mockResolvedValue({
        value: [],
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), ["mail", "delta", "--max-page-size", "100"]);

      const opts = vi.mocked(graph.get).mock.calls[0][1] as {
        headers?: Record<string, string>;
      };
      expect(opts?.headers?.["Prefer"]).toBe("odata.maxpagesize=100");
    });
  });

  // --- --change-type ---

  describe("--change-type", () => {
    it("passes changeType to initial URL query params", async () => {
      noStateFile();
      vi.mocked(graph.get).mockResolvedValue({
        value: [],
        "@odata.deltaLink": FAKE_DELTA_LINK,
      });

      await parse(makeProgram(), ["mail", "delta", "--change-type", "created"]);

      const url = vi.mocked(graph.get).mock.calls[0][0] as string;
      expect(decodeURIComponent(url)).toContain("changeType=created");
    });
  });
});
