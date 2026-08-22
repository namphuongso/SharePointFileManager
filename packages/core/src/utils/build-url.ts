import type { RestRequestOptions } from "../types/rest";

/**
 * Ghép `/_api/{path}` và query OData.
 * encodeURIComponent: khoảng trắng = `%20`. URLSearchParams dùng `+` — SharePoint OData coi `+` là toán tử.
 * nextLink thì dùng nguyên URL.
 */
export function buildRestUrl(apiUrl: (path: string) => string, options: RestRequestOptions): string {
  if (options.absoluteUrl) return options.absoluteUrl;
  const base = apiUrl(options.path);
  if (!options.query) return base;
  const qs = Object.entries(options.query)
    .filter(([, value]) => value !== undefined)
    .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(String(value))}`)
    .join("&");
  return qs ? `${base}?${qs}` : base;
}
