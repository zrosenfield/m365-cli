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

// --- import mocked modules ---

import { graph, GraphError } from "../../src/lib/graph.js";

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

// --- test suite ---

describe("mail commands", () => {
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
    vi.mocked(graph.post).mockReset();
    vi.mocked(graph.delete).mockReset();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  // --- mail list ---

  describe("mail list", () => {
    it("defaults to inbox with top=25", async () => {
      vi.mocked(graph.get).mockResolvedValue({ value: [] });

      await parse(makeProgram(), ["mail", "list"]);

      const url = vi.mocked(graph.get).mock.calls[0][0] as string;
      expect(url).toContain("/me/mailFolders/inbox/messages");
      // URLSearchParams encodes $ as %24
      expect(url).toContain("%24top=25");
    });

    it("uses well-known folder name (sentitems)", async () => {
      vi.mocked(graph.get).mockResolvedValue({ value: [] });

      await parse(makeProgram(), ["mail", "list", "--folder", "sentItems"]);

      const url = vi.mocked(graph.get).mock.calls[0][0] as string;
      expect(url).toContain("/me/mailFolders/sentitems/messages");
    });

    it("respects --top flag", async () => {
      vi.mocked(graph.get).mockResolvedValue({ value: [] });

      await parse(makeProgram(), ["mail", "list", "--top", "5"]);

      const url = vi.mocked(graph.get).mock.calls[0][0] as string;
      expect(url).toContain("%24top=5");
    });

    it("outputs messages array under data key", async () => {
      const messages = [{ id: "m1", subject: "Hello" }];
      vi.mocked(graph.get).mockResolvedValue({ value: messages });

      await parse(makeProgram(), ["mail", "list"]);

      expect(JSON.parse(stdoutData).data).toEqual(messages);
    });

    it("forwards --select to query params", async () => {
      vi.mocked(graph.get).mockResolvedValue({ value: [] });

      await parse(makeProgram(), ["mail", "list", "--select", "id,subject"]);

      const url = vi.mocked(graph.get).mock.calls[0][0] as string;
      expect(url).toContain("%24select=id%2Csubject");
    });

    it("outputs error to stderr on graph failure", async () => {
      vi.mocked(graph.get).mockRejectedValue(
        new GraphError(401, "Unauthorized", "Token expired")
      );

      await expect(parse(makeProgram(), ["mail", "list"])).rejects.toThrow(
        "process.exit(1)"
      );

      expect(JSON.parse(stderrData).error.status).toBe(401);
    });
  });

  // --- mail get ---

  describe("mail get", () => {
    it("calls graph.get with the message path", async () => {
      const msg = { id: "msg1", subject: "Test" };
      vi.mocked(graph.get).mockResolvedValue(msg);

      await parse(makeProgram(), ["mail", "get", "msg1"]);

      expect(graph.get).toHaveBeenCalledWith("/me/messages/msg1");
      expect(JSON.parse(stdoutData).data).toEqual(msg);
    });
  });

  // --- mail send ---

  describe("mail send", () => {
    it("calls graph.post with correct message body", async () => {
      vi.mocked(graph.post).mockResolvedValue(undefined);

      await parse(makeProgram(), [
        "mail",
        "send",
        "--to",
        "user@example.com",
        "--subject",
        "Hello",
        "--body",
        "World",
      ]);

      expect(graph.post).toHaveBeenCalledWith(
        "/me/sendMail",
        expect.objectContaining({
          message: expect.objectContaining({
            subject: "Hello",
            toRecipients: [{ emailAddress: { address: "user@example.com" } }],
            body: { contentType: "Text", content: "World" },
          }),
          saveToSentItems: true,
        })
      );
      expect(JSON.parse(stdoutData).data.message).toBe("Mail sent.");
    });

    it("uses HTML contentType when --html flag is set", async () => {
      vi.mocked(graph.post).mockResolvedValue(undefined);

      await parse(makeProgram(), [
        "mail",
        "send",
        "--to",
        "a@b.com",
        "--subject",
        "s",
        "--body",
        "<p>hi</p>",
        "--html",
      ]);

      const body = vi.mocked(graph.post).mock.calls[0][1] as {
        message: { body: { contentType: string } };
      };
      expect(body.message.body.contentType).toBe("HTML");
    });

    it("rejects invalid email addresses", async () => {
      await expect(
        parse(makeProgram(), [
          "mail",
          "send",
          "--to",
          "notanemail",
          "--subject",
          "s",
          "--body",
          "b",
        ])
      ).rejects.toThrow("process.exit(1)");

      expect(stderrData).toContain("Invalid email address");
    });

    it("includes CC recipients when --cc is provided", async () => {
      vi.mocked(graph.post).mockResolvedValue(undefined);

      await parse(makeProgram(), [
        "mail",
        "send",
        "--to",
        "to@example.com",
        "--subject",
        "s",
        "--body",
        "b",
        "--cc",
        "cc1@example.com,cc2@example.com",
      ]);

      const body = vi.mocked(graph.post).mock.calls[0][1] as {
        message: { ccRecipients: { emailAddress: { address: string } }[] };
      };
      expect(body.message.ccRecipients).toHaveLength(2);
    });
  });

  // --- mail delete ---

  describe("mail delete", () => {
    it("calls graph.delete and outputs confirmation", async () => {
      vi.mocked(graph.delete).mockResolvedValue(undefined);

      await parse(makeProgram(), ["mail", "delete", "msg99"]);

      expect(graph.delete).toHaveBeenCalledWith("/me/messages/msg99");
      expect(JSON.parse(stdoutData).data.message).toContain("msg99");
    });
  });

  // --- mail folders list ---

  describe("mail folders list", () => {
    it("calls graph.get for mailFolders", async () => {
      vi.mocked(graph.get).mockResolvedValue({ value: [{ id: "f1" }] });

      await parse(makeProgram(), ["mail", "folders", "list"]);

      expect(graph.get).toHaveBeenCalledWith(
        "/me/mailFolders?includeHiddenFolders=false"
      );
    });
  });
});
