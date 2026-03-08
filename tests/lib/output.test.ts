import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { outputData, outputError } from "../../src/lib/output.js";
import { GraphError } from "../../src/lib/graph.js";

describe("outputData", () => {
  let captured: string;

  beforeEach(() => {
    captured = "";
    vi.spyOn(process.stdout, "write").mockImplementation((data) => {
      captured += String(data);
      return true;
    });
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("wraps value in { data } envelope", () => {
    outputData([1, 2, 3]);
    expect(JSON.parse(captured)).toEqual({ data: [1, 2, 3] });
  });

  it("writes raw JSON when raw=true", () => {
    outputData({ foo: "bar" }, true);
    expect(JSON.parse(captured)).toEqual({ foo: "bar" });
  });

  it("pretty-prints by default (has newline)", () => {
    outputData("hello");
    expect(captured).toContain("\n");
  });

  it("no newline in raw mode", () => {
    outputData("hello", true);
    expect(captured).not.toContain("\n");
  });
});

describe("outputError", () => {
  let captured: string;

  beforeEach(() => {
    captured = "";
    vi.spyOn(process.stderr, "write").mockImplementation((data) => {
      captured += String(data);
      return true;
    });
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("formats a GraphError with status and code", () => {
    outputError(new GraphError(404, "ItemNotFound", "Item not found"));
    const out = JSON.parse(captured);
    expect(out.error).toEqual({
      code: "ItemNotFound",
      message: "Item not found",
      status: 404,
    });
  });

  it("formats a plain Error", () => {
    outputError(new Error("something went wrong"));
    const out = JSON.parse(captured);
    expect(out.error).toEqual({ code: "Error", message: "something went wrong" });
  });

  it("formats an unknown throw value", () => {
    outputError("oops");
    const out = JSON.parse(captured);
    expect(out.error).toEqual({ code: "UnknownError", message: "oops" });
  });

  it("writes raw JSON when raw=true", () => {
    outputError(new Error("raw"), true);
    const out = JSON.parse(captured);
    expect(out.error.code).toBe("Error");
    expect(captured).not.toContain("\n");
  });
});
