import { describe, it, expect } from "vitest";
import { validateId, GraphError } from "../../src/lib/graph.js";

describe("validateId", () => {
  it("accepts plain IDs", () => {
    expect(() => validateId("abc123")).not.toThrow();
    expect(() => validateId("site-id_with.chars")).not.toThrow();
    expect(() => validateId("ABC123XYZ")).not.toThrow();
  });

  it("rejects IDs containing /", () => {
    expect(() => validateId("a/b")).toThrow("Invalid ID");
    expect(() => validateId("/etc/passwd")).toThrow("Invalid ID");
  });

  it("rejects IDs containing \\", () => {
    expect(() => validateId("a\\b")).toThrow("Invalid ID");
  });

  it("rejects IDs containing ..", () => {
    expect(() => validateId("a..b")).toThrow("Invalid ID");
    expect(() => validateId("../etc")).toThrow("Invalid ID");
  });

  it("uses the name parameter in the error message", () => {
    expect(() => validateId("bad/id", "drive ID")).toThrow("Invalid drive ID");
  });
});

describe("GraphError", () => {
  it("stores status, code, and message", () => {
    const err = new GraphError(403, "Forbidden", "Access denied");
    expect(err.status).toBe(403);
    expect(err.code).toBe("Forbidden");
    expect(err.message).toBe("Access denied");
    expect(err.name).toBe("GraphError");
  });

  it("is an instance of Error", () => {
    expect(new GraphError(500, "ServerError", "oops")).toBeInstanceOf(Error);
  });

  it("is an instance of GraphError", () => {
    expect(new GraphError(404, "NotFound", "not found")).toBeInstanceOf(GraphError);
  });
});
