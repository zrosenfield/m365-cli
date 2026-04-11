import { graph } from "./graph.js";

const siteIdCache = new Map<string, string>();

/** Reset the site-ID cache — useful in tests. */
export function clearSiteIdCache(): void {
  siteIdCache.clear();
}

/**
 * Resolve a --site value to a bare site ID.
 *
 * If the value looks like a SharePoint URL (starts with https:// or contains
 * .sharepoint.com) it is resolved via GET /sites/{hostname}:{path} and the
 * returned id is cached for the lifetime of the process.  Otherwise the value
 * is assumed to already be a site ID and is returned unchanged.
 */
export async function resolveSiteId(value: string): Promise<string> {
  // A compound site ID has the form  "hostname,guid1,guid2" — it contains
  // ".sharepoint.com" but never a "/" after the hostname.  We require either
  // an explicit https:// scheme or ".sharepoint.com/" (with trailing slash,
  // indicating a path) so that compound IDs are not mistaken for URLs.
  const isUrl = value.startsWith("https://") || value.startsWith("http://")
    || /\.sharepoint\.com\//.test(value);
  if (!isUrl) return value;

  if (siteIdCache.has(value)) return siteIdCache.get(value)!;

  const normalized = value.startsWith("https://") ? value : `https://${value}`;
  const url = new URL(normalized);
  const hostname = url.hostname;
  const sitePath = url.pathname; // e.g. /sites/matter-henderson

  const result = await graph.get<{ id: string }>(`/sites/${hostname}:${sitePath}`);
  if (!result.id) throw new Error(`Could not resolve site URL to an ID: ${value}`);

  siteIdCache.set(value, result.id);
  return result.id;
}

export interface DriveItemMeta {
  id: string;
  name?: string;
  "@microsoft.graph.downloadUrl"?: string;
}

/**
 * Resolve a remote path to a DriveItem via GET /drives/{driveId}/root:/{path}.
 *
 * The path may or may not start with a leading slash — one is added if absent.
 */
export async function resolveItemByPath(driveId: string, itemPath: string): Promise<DriveItemMeta> {
  const normalized = itemPath.startsWith("/") ? itemPath : `/${itemPath}`;
  return graph.get<DriveItemMeta>(`/drives/${driveId}/root:${normalized}`);
}
