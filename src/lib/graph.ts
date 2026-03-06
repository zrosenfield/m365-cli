import fetch, { Response, RequestInit } from "node-fetch";
import { getAccessToken } from "./auth.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const GRAPH_HOST = "graph.microsoft.com";

export class GraphError extends Error {
  constructor(
    public readonly status: number,
    public readonly code: string,
    message: string
  ) {
    super(message);
    this.name = "GraphError";
  }
}

async function handleResponse<T>(res: Response): Promise<T> {
  if (res.status === 204) return undefined as unknown as T;

  const contentType = res.headers.get("content-type") ?? "";
  const isJson = contentType.includes("application/json");

  if (!res.ok) {
    let errCode = "UnknownError";
    let errMsg = `HTTP ${res.status} ${res.statusText}`;
    if (isJson) {
      const body = (await res.json()) as { error?: { code?: string; message?: string } };
      errCode = body.error?.code ?? errCode;
      errMsg = body.error?.message ?? errMsg;
    }
    throw new GraphError(res.status, errCode, errMsg);
  }

  if (isJson) return res.json() as Promise<T>;
  return res.buffer() as unknown as Promise<T>;
}

export function validateId(value: string, name = "ID"): void {
  if (value.includes("/") || value.includes("\\") || value.includes("..")) {
    throw new Error(`Invalid ${name}: must not contain path separators.`);
  }
}

async function request<T>(
  method: string,
  path: string,
  options: {
    body?: unknown;
    rawBody?: Buffer | string;
    contentType?: string;
    headers?: Record<string, string>;
  } = {}
): Promise<T> {
  const token = await getAccessToken();
  let url: string;
  if (path.startsWith("https://")) {
    const parsed = new URL(path);
    if (parsed.hostname !== GRAPH_HOST) {
      throw new Error(`Refused to send credentials to non-Graph host: ${parsed.hostname}`);
    }
    url = path;
  } else {
    url = `${GRAPH_BASE}${path}`;
  }

  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    Accept: "application/json",
    ...options.headers,
  };

  let body: RequestInit["body"];
  if (options.rawBody !== undefined) {
    body = options.rawBody;
    headers["Content-Type"] = options.contentType ?? "application/octet-stream";
  } else if (options.body !== undefined) {
    body = JSON.stringify(options.body);
    headers["Content-Type"] = "application/json";
  }

  const res = await fetch(url, { method, headers, body });
  return handleResponse<T>(res);
}

export const graph = {
  get: <T>(path: string) => request<T>("GET", path),
  post: <T>(path: string, body: unknown) => request<T>("POST", path, { body }),
  patch: <T>(path: string, body: unknown) => request<T>("PATCH", path, { body }),
  put: <T>(path: string, body: unknown) => request<T>("PUT", path, { body }),
  delete: <T>(path: string) => request<T>("DELETE", path),
  upload: <T>(path: string, data: Buffer | string, contentType?: string) =>
    request<T>("PUT", path, { rawBody: data, contentType }),
};

export type { Response };
