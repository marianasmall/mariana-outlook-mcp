const GRAPH = "https://graph.microsoft.com/v1.0";

export interface GraphOptions {
  method?: "GET" | "POST" | "PATCH" | "DELETE";
  body?: unknown;
  headers?: Record<string, string>;
}

/**
 * Minimal Microsoft Graph client over global fetch. Throws with Graph's own
 * error message on non-2xx. Returns parsed JSON (or null for 204/empty).
 */
export async function graphFetch(
  accessToken: string,
  path: string,
  opts: GraphOptions = {}
): Promise<any> {
  const res = await fetch(`${GRAPH}${path}`, {
    method: opts.method || "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(opts.body !== undefined ? { "Content-Type": "application/json" } : {}),
      ...(opts.headers || {}),
    },
    body: opts.body !== undefined ? JSON.stringify(opts.body) : undefined,
  });

  if (res.status === 204) return null;
  const text = await res.text();
  const data = text ? JSON.parse(text) : null;
  if (!res.ok) {
    const msg = data?.error?.message || `${res.status} ${res.statusText}`;
    throw new Error(`Graph ${opts.method || "GET"} ${path.split("?")[0]} failed: ${msg}`);
  }
  return data;
}
