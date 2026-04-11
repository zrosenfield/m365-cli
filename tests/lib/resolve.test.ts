import { describe, it, expect, vi, beforeEach } from "vitest";

vi.mock("../../src/lib/graph.js", () => ({
  graph: {
    get: vi.fn(),
  },
}));

import { graph } from "../../src/lib/graph.js";
import { resolveSiteId, resolveItemByPath, clearSiteIdCache } from "../../src/lib/resolve.js";

describe("resolveSiteId", () => {
  beforeEach(() => {
    vi.mocked(graph.get).mockReset();
    clearSiteIdCache();
  });

  it("returns the value unchanged when it is a plain ID (no URL)", async () => {
    const id = "site-id-123";
    const result = await resolveSiteId(id);
    expect(result).toBe(id);
    expect(graph.get).not.toHaveBeenCalled();
  });

  it("resolves an https:// URL to the site ID via Graph", async () => {
    vi.mocked(graph.get).mockResolvedValue({ id: "resolved-site-id" });

    const result = await resolveSiteId("https://cmalaw.sharepoint.com/sites/matter-henderson");

    expect(graph.get).toHaveBeenCalledWith(
      "/sites/cmalaw.sharepoint.com:/sites/matter-henderson"
    );
    expect(result).toBe("resolved-site-id");
  });

  it("detects a .sharepoint.com host with a path (no https:// prefix)", async () => {
    vi.mocked(graph.get).mockResolvedValue({ id: "sp-site-id" });

    const result = await resolveSiteId("example.sharepoint.com/sites/foo");

    expect(graph.get).toHaveBeenCalledWith("/sites/example.sharepoint.com:/sites/foo");
    expect(result).toBe("sp-site-id");
  });

  it("does NOT treat a compound site ID (hostname,guid1,guid2) as a URL", async () => {
    const compoundId = "contoso.sharepoint.com,abc-123,def-456";
    const result = await resolveSiteId(compoundId);
    expect(result).toBe(compoundId);
    expect(graph.get).not.toHaveBeenCalled();
  });

  it("caches resolved site IDs — Graph is called only once for the same URL", async () => {
    vi.mocked(graph.get).mockResolvedValue({ id: "cached-id" });
    const url = "https://contoso.sharepoint.com/sites/test";

    const first = await resolveSiteId(url);
    const second = await resolveSiteId(url);

    expect(first).toBe("cached-id");
    expect(second).toBe("cached-id");
    expect(graph.get).toHaveBeenCalledTimes(1);
  });

  it("throws when Graph returns no id field", async () => {
    vi.mocked(graph.get).mockResolvedValue({});

    await expect(
      resolveSiteId("https://bad.sharepoint.com/sites/x")
    ).rejects.toThrow("Could not resolve site URL");
  });
});

describe("resolveItemByPath", () => {
  beforeEach(() => {
    vi.mocked(graph.get).mockReset();
  });

  it("calls Graph with the path-based drive endpoint", async () => {
    const item = { id: "item-abc", name: "report.docx" };
    vi.mocked(graph.get).mockResolvedValue(item);

    const result = await resolveItemByPath("drive1", "/Documents/report.docx");

    expect(graph.get).toHaveBeenCalledWith("/drives/drive1/root:/Documents/report.docx");
    expect(result).toEqual(item);
  });

  it("adds a leading slash when the path omits it", async () => {
    vi.mocked(graph.get).mockResolvedValue({ id: "x" });

    await resolveItemByPath("drive1", "Documents/file.txt");

    expect(graph.get).toHaveBeenCalledWith("/drives/drive1/root:/Documents/file.txt");
  });
});
