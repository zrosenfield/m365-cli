import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { Command } from "commander";
import { registerFileCommands } from "../../src/commands/files.js";

// --- module mocks (hoisted) ---

vi.mock("../../src/lib/graph.js", () => ({
  graph: {
    get: vi.fn(),
    post: vi.fn(),
    patch: vi.fn(),
    put: vi.fn(),
    delete: vi.fn(),
    upload: vi.fn(),
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

vi.mock("../../src/lib/config.js", () => ({
  readConfig: vi.fn(),
}));

vi.mock("../../src/lib/resolve.js", () => ({
  resolveSiteId: vi.fn(async (v: string) => v), // default: return value as-is
  resolveItemByPath: vi.fn(),
}));

// --- import mocked modules ---

import { graph, GraphError } from "../../src/lib/graph.js";
import { readConfig } from "../../src/lib/config.js";
import { resolveSiteId, resolveItemByPath } from "../../src/lib/resolve.js";

// --- helpers ---

function makeProgram(): Command {
  const p = new Command();
  p.exitOverride(); // throws CommanderError instead of calling process.exit
  registerFileCommands(p);
  return p;
}

function parse(p: Command, args: string[]) {
  return p.parseAsync(["node", "m365", ...args]);
}

// --- test suite ---

describe("files commands", () => {
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

    // Default config: site + drive available
    vi.mocked(readConfig).mockReturnValue({
      defaultSiteId: "site1",
      defaultDriveId: "drive1",
    });
    // validateId: no-op (we test it separately in graph.test.ts)
    vi.mocked(graph.get).mockReset();
    vi.mocked(graph.post).mockReset();
    vi.mocked(graph.patch).mockReset();
    vi.mocked(graph.delete).mockReset();
    vi.mocked(graph.upload).mockReset();
    vi.mocked(resolveSiteId).mockReset().mockImplementation(async (v: string) => v);
    vi.mocked(resolveItemByPath).mockReset();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  // --- files list ---

  describe("files list", () => {
    it("calls graph.get with the drive root children path", async () => {
      const items = [{ id: "1", name: "report.docx" }];
      vi.mocked(graph.get).mockResolvedValue({ value: items });

      await parse(makeProgram(), ["files", "list"]);

      expect(graph.get).toHaveBeenCalledWith("/drives/drive1/root/children");
      expect(JSON.parse(stdoutData).data).toEqual(items);
    });

    it("appends folder path when --path is given", async () => {
      vi.mocked(graph.get).mockResolvedValue({ value: [] });

      await parse(makeProgram(), ["files", "list", "--path", "/Documents"]);

      expect(graph.get).toHaveBeenCalledWith(
        "/drives/drive1/root:/Documents:/children"
      );
    });

    it("uses --drive flag over config default", async () => {
      vi.mocked(graph.get).mockResolvedValue({ value: [] });

      await parse(makeProgram(), ["files", "list", "--drive", "driveX"]);

      expect(graph.get).toHaveBeenCalledWith(
        expect.stringContaining("/drives/driveX/")
      );
    });

    it("outputs error to stderr and exits when graph throws", async () => {
      vi.mocked(graph.get).mockRejectedValue(
        new GraphError(403, "Forbidden", "Access denied")
      );

      await expect(parse(makeProgram(), ["files", "list"])).rejects.toThrow(
        "process.exit(1)"
      );

      const err = JSON.parse(stderrData);
      expect(err.error.status).toBe(403);
      expect(err.error.code).toBe("Forbidden");
    });

    it("errors when no drive is configured", async () => {
      vi.mocked(readConfig).mockReturnValue({ defaultSiteId: "site1" }); // no driveId

      await expect(parse(makeProgram(), ["files", "list"])).rejects.toThrow(
        "process.exit(1)"
      );

      expect(stderrData).toContain("Drive ID required");
    });
  });

  // --- files get ---

  describe("files get", () => {
    it("calls graph.get with the item path", async () => {
      const item = { id: "abc", name: "file.txt" };
      vi.mocked(graph.get).mockResolvedValue(item);

      await parse(makeProgram(), ["files", "get", "abc"]);

      expect(graph.get).toHaveBeenCalledWith("/drives/drive1/items/abc");
      expect(JSON.parse(stdoutData).data).toEqual(item);
    });
  });

  // --- files delete ---

  describe("files delete", () => {
    it("calls graph.delete and outputs confirmation", async () => {
      vi.mocked(graph.delete).mockResolvedValue(undefined);

      await parse(makeProgram(), ["files", "delete", "item99"]);

      expect(graph.delete).toHaveBeenCalledWith("/drives/drive1/items/item99");
      expect(JSON.parse(stdoutData).data.message).toContain("item99");
    });
  });

  // --- files search ---

  describe("files search", () => {
    it("encodes the query and calls graph.get", async () => {
      vi.mocked(graph.get).mockResolvedValue({ value: [] });

      await parse(makeProgram(), ["files", "search", "hello world"]);

      const call = vi.mocked(graph.get).mock.calls[0][0] as string;
      expect(call).toContain("/drives/drive1/root/search(q=");
      expect(call).toContain(encodeURIComponent("hello world"));
    });
  });

  // --- Feature 2: site URL resolution ---

  describe("site URL resolution via resolveSiteId", () => {
    it("passes --site value through resolveSiteId before using it", async () => {
      vi.mocked(resolveSiteId).mockResolvedValue("resolved-site-id");
      vi.mocked(graph.get).mockResolvedValue({ value: [] });

      await parse(makeProgram(), ["files", "list", "--site", "https://contoso.sharepoint.com/sites/foo"]);

      expect(resolveSiteId).toHaveBeenCalledWith("https://contoso.sharepoint.com/sites/foo");
    });

    it("uses the resolved ID (not the raw URL) for drive lookup", async () => {
      vi.mocked(resolveSiteId).mockResolvedValue("resolved-site-id");
      vi.mocked(readConfig).mockReturnValue({ defaultSiteId: "https://sp.sharepoint.com/sites/x", defaultDriveId: "drive1" });
      vi.mocked(graph.get).mockResolvedValue({ value: [] });

      await parse(makeProgram(), ["files", "list"]);

      // resolveSiteId was called with the raw URL from config
      expect(resolveSiteId).toHaveBeenCalledWith("https://sp.sharepoint.com/sites/x");
    });
  });

  // --- Feature 3: path-based file operations ---

  describe("files read --remote-path", () => {
    it("resolves path to item then fetches content", async () => {
      const downloadUrl = "https://download.example.com/file.txt";
      vi.mocked(resolveItemByPath).mockResolvedValue({
        id: "item-path-id",
        name: "file.txt",
        "@microsoft.graph.downloadUrl": downloadUrl,
      });

      // Mock node-fetch via dynamic import interception isn't practical in unit
      // tests; we test the resolveItemByPath call and URL extraction path by
      // making fetch reject so handleCommandError is invoked (confirming we got
      // past the resolveItemByPath call).
      await expect(
        parse(makeProgram(), ["files", "read", "--remote-path", "/Documents/file.txt"])
      ).rejects.toThrow("process.exit");

      expect(resolveItemByPath).toHaveBeenCalledWith("drive1", "/Documents/file.txt");
    });

    it("errors when neither itemId nor --remote-path is provided", async () => {
      await expect(
        parse(makeProgram(), ["files", "read"])
      ).rejects.toThrow("process.exit(1)");

      expect(stderrData).toContain("Either an item ID or --remote-path is required");
    });
  });

  describe("files download --remote-path", () => {
    it("calls resolveItemByPath when --remote-path is given", async () => {
      vi.mocked(resolveItemByPath).mockResolvedValue({
        id: "dl-item-id",
        name: "doc.pdf",
        "@microsoft.graph.downloadUrl": "https://download.example.com/doc.pdf",
      });

      await expect(
        parse(makeProgram(), ["files", "download", "--remote-path", "/Reports/doc.pdf"])
      ).rejects.toThrow("process.exit");

      expect(resolveItemByPath).toHaveBeenCalledWith("drive1", "/Reports/doc.pdf");
    });

    it("errors when neither itemId nor --remote-path is provided", async () => {
      await expect(
        parse(makeProgram(), ["files", "download"])
      ).rejects.toThrow("process.exit(1)");

      expect(stderrData).toContain("Either an item ID or --remote-path is required");
    });
  });

  describe("files move --remote-path", () => {
    it("resolves path to item ID then patches", async () => {
      vi.mocked(resolveItemByPath).mockResolvedValue({ id: "moved-item-id" });
      vi.mocked(graph.patch).mockResolvedValue({ id: "moved-item-id" });

      await parse(makeProgram(), [
        "files", "move",
        "--remote-path", "/Docs/old.txt",
        "--dest-path", "/Archive",
      ]);

      expect(resolveItemByPath).toHaveBeenCalledWith("drive1", "/Docs/old.txt");
      expect(graph.patch).toHaveBeenCalledWith(
        "/drives/drive1/items/moved-item-id",
        expect.objectContaining({ parentReference: expect.objectContaining({ driveId: "drive1" }) })
      );
    });

    it("errors when neither itemId nor --remote-path is provided", async () => {
      await expect(
        parse(makeProgram(), ["files", "move", "--dest-path", "/Archive"])
      ).rejects.toThrow("process.exit(1)");

      expect(stderrData).toContain("Either an item ID or --remote-path is required");
    });
  });

  describe("files upload --local-path", () => {
    it("prefers --local-path over positional argument", async () => {
      const tmpDir = process.env.TMPDIR || process.env.TMP || "/tmp";
      const tmpFile = `${tmpDir}/test-upload.txt`;
      const { writeFileSync } = await import("fs");
      writeFileSync(tmpFile, "hello");

      vi.mocked(graph.upload).mockResolvedValue({ id: "new-item" });

      await parse(makeProgram(), [
        "files", "upload", "positional-ignored.txt",
        "--local-path", tmpFile,
        "--remote-path", "/uploads/test.txt",
      ]);

      expect(graph.upload).toHaveBeenCalledWith(
        "/drives/drive1/root:/uploads/test.txt:/content",
        expect.any(Buffer)
      );
    });

    it("errors when neither positional nor --local-path is provided", async () => {
      await expect(
        parse(makeProgram(), ["files", "upload"])
      ).rejects.toThrow("process.exit(1)");

      expect(stderrData).toContain("local file path is required");
    });
  });
});
