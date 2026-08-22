import type { TokenProvider } from "../auth/token-provider";
import { mapRestError } from "../errors/map-rest-error";
import { SharePointError, SharePointErrorCode } from "../errors/sharepoint-error";

export interface RestRequestOptions {
  path: string;
  method?: string;
  query?: Record<string, string | number | boolean | undefined>;
  headers?: Record<string, string>;
  signal?: AbortSignal;
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

/**
 * SharePoint REST client — GET only (list / browse).
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service
 */
export class SharePointRestClient {
  constructor(
    private readonly options: {
      siteUrl: string;
      tokenProvider: TokenProvider;
      scopes: string[];
      fetchImpl?: typeof fetch;
    },
  ) {}

  get siteUrl(): string {
    return this.options.siteUrl.replace(/\/$/, "");
  }

  apiUrl(path: string): string {
    const cleaned = path.replace(/^\/+/, "").replace(/^_api\/?/i, "");
    return `${this.siteUrl}/_api/${cleaned}`;
  }

  async get<T>(path: string, options: Omit<RestRequestOptions, "path" | "method"> = {}): Promise<T> {
    return this.request<T>({ ...options, path, method: "GET" });
  }

  async request<T>(options: RestRequestOptions): Promise<T> {
    const fetchImpl = this.options.fetchImpl ?? fetch;
    const method = options.method ?? "GET";
    let unauthorizedRetried = false;
    let throttleAttempt = 0;

    while (true) {
      if (options.signal?.aborted) {
        throw new SharePointError({
          code: SharePointErrorCode.Cancelled,
          message: "Request was cancelled",
        });
      }

      const url = this.buildUrl(options);
      const headers = new Headers(options.headers);
      if (!headers.has("Accept")) {
        headers.set("Accept", "application/json;odata=nometadata");
      }

      const token = await this.options.tokenProvider.getAccessToken({
        scopes: this.options.scopes,
        forceRefresh: unauthorizedRetried,
      });
      headers.set("Authorization", `Bearer ${token}`);

      let response: Response;
      try {
        response = await fetchImpl(url, { method, headers, signal: options.signal });
      } catch (error) {
        if (options.signal?.aborted) {
          throw new SharePointError({
            code: SharePointErrorCode.Cancelled,
            message: "Request was cancelled",
            cause: error,
          });
        }
        throw new SharePointError({
          code: SharePointErrorCode.NetworkError,
          message: "SharePoint REST request failed",
          cause: error,
        });
      }

      if (response.status === 401 && !unauthorizedRetried) {
        unauthorizedRetried = true;
        continue;
      }

      if (response.status === 429 && throttleAttempt < 3) {
        throttleAttempt += 1;
        const retryAfter = response.headers.get("Retry-After");
        const waitMs = retryAfter ? Number(retryAfter) * 1000 || 2000 : 1000 * throttleAttempt;
        await sleep(waitMs);
        continue;
      }

      if (!response.ok) {
        const body = await safeJson(response);
        throw mapRestError({
          status: response.status,
          body,
          retryAfter: response.headers.get("Retry-After"),
        });
      }

      if (response.status === 204 || response.headers.get("Content-Length") === "0") {
        return undefined as T;
      }

      const text = await response.text();
      if (!text) return undefined as T;
      try {
        return JSON.parse(text) as T;
      } catch {
        return text as unknown as T;
      }
    }
  }

  private buildUrl(options: RestRequestOptions): string {
    const base = this.apiUrl(options.path);
    if (!options.query) return base;
    const params = new URLSearchParams();
    for (const [key, value] of Object.entries(options.query)) {
      if (value === undefined) continue;
      params.set(key, String(value));
    }
    const qs = params.toString();
    return qs ? `${base}?${qs}` : base;
  }
}

async function safeJson(response: Response): Promise<unknown> {
  try {
    return await response.json();
  } catch {
    return undefined;
  }
}
