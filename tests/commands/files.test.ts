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

// --- import mocked modules ---

import { graph, GraphError } from "../../src/lib/graph.js";
import { readConfig } from "../../src/lib/config.js";

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
});
