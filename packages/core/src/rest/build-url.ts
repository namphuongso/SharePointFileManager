import type { RestRequestOptions } from "../types/rest";

/** Ghép `/_api/{path}` và query OData (`$select`, `$filter`, …). */
export function buildRestUrl(apiUrl: (path: string) => string, options: RestRequestOptions): string {
  const base = apiUrl(options.path);
  if (!options.query) return base;
  const params = new URLSearchParams();
  for (const [key, value] of Object.entries(options.query)) {
    if (value === undefined) continue;
    params.set(key, String(value));
  }
  const qs = params.toString();
  return qs ? `${base}?${qs}` : base;
}
