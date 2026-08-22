import { normalizeSiteUrl } from "./site-url";

/** Bỏ slash đầu và prefix `_api` để ghép vào `{site}/_api/{path}`. */
export function normalizeApiPath(path: string): string {
  return path.replace(/^\/+/, "").replace(/^_api\/?/i, "");
}

export function buildApiUrl(siteUrl: string, path: string): string {
  return `${normalizeSiteUrl(siteUrl)}/_api/${normalizeApiPath(path)}`;
}
