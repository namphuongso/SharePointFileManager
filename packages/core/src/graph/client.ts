import type { TokenProvider } from "../auth/token-provider";
import { mapGraphError } from "../errors/map-graph-error";
import { SharePointError, SharePointErrorCode } from "../errors/sharepoint-error";

export interface GraphRequestOptions {
  path: string;
  method?: string;
  query?: Record<string, string | number | boolean | undefined>;
  body?: unknown;
  headers?: Record<string, string>;
  signal?: AbortSignal;
  /** Skip JSON parsing and return the raw Response. */
  raw?: boolean;
  /** Do not attach a bearer token (pre-authenticated download/upload URLs). */
  skipAuth?: boolean;
  /** Treat this URL as absolute instead of graphBaseUrl + path. */
  absoluteUrl?: boolean;
  timeoutMs?: number;
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

export class GraphClient {
  constructor(
    private readonly options: {
      baseUrl: string;
      tokenProvider: TokenProvider;
      scopes: string[];
      fetchImpl?: typeof fetch;
    },
  ) {}

  /** Build an endpoint against the configured Graph cloud v1.0 API. */
  apiUrl(path: string): string {
    const configuredBase = this.options.baseUrl.replace(/\/v1\.0\/?$/i, "");
    return joinUrl(`${configuredBase}/v1.0`, path);
  }

  async request<T>(options: GraphRequestOptions): Promise<T> {
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

      if (!options.skipAuth) {
        const token = await this.options.tokenProvider.getAccessToken({
          scopes: this.options.scopes,
          forceRefresh: unauthorizedRetried,
        });
        headers.set("Authorization", `Bearer ${token}`);
      }

      const init: RequestInit = {
        method,
        headers,
        signal: options.signal,
      };

      if (options.body instanceof Uint8Array || options.body instanceof ArrayBuffer || options.body instanceof Blob) {
        init.body = options.body as BodyInit;
      } else if (options.body !== undefined) {
        if (!headers.has("Content-Type")) {
          headers.set("Content-Type", "application/json");
        }
        init.body = JSON.stringify(options.body);
      }

      let response: Response;
      const timeout = options.timeoutMs && options.timeoutMs > 0 ? new AbortController() : undefined;
      let timeoutId: ReturnType<typeof setTimeout> | undefined;
      if (timeout) {
        timeoutId = setTimeout(() => timeout.abort(), options.timeoutMs);
      }
      try {
        const signal = timeout ? combineSignals(options.signal, timeout.signal) : options.signal;
        response = await fetchImpl(url, { ...init, signal });
      } catch (error) {
        if (timeoutId) clearTimeout(timeoutId);
        if (timeout?.signal.aborted && !options.signal?.aborted) {
          throw new SharePointError({ code: SharePointErrorCode.NetworkError, message: "Graph request timed out", cause: error });
        }
        if (options.signal?.aborted) {
          throw new SharePointError({
            code: SharePointErrorCode.Cancelled,
            message: "Request was cancelled",
            cause: error,
          });
        }
        throw new SharePointError({
          code: SharePointErrorCode.NetworkError,
          message: error instanceof Error ? error.message : "Network request failed",
          cause: error,
        });
      }
      if (timeoutId) clearTimeout(timeoutId);

      if (response.status === 401 && !unauthorizedRetried && !options.skipAuth) {
        unauthorizedRetried = true;
        continue;
      }

      if (response.status === 429 && throttleAttempt < 3) {
        const retryAfter = response.headers.get("Retry-After");
        const waitMs = parseRetryAfterOrBackoff(retryAfter, throttleAttempt);
        throttleAttempt += 1;
        await sleep(waitMs);
        continue;
      }

      if (!response.ok) {
        const body = await safeJson(response);
        throw mapGraphError({
          status: response.status,
          body,
          retryAfter: response.headers.get("Retry-After"),
        });
      }

      if (options.raw) {
        return response as T;
      }

      if (response.status === 204) {
        return undefined as T;
      }

      const contentType = response.headers.get("Content-Type") ?? "";
      if (contentType.includes("application/json") || contentType.includes("+json")) {
        return (await response.json()) as T;
      }

      return undefined as T;
    }
  }

  async get<T>(path: string, options: Omit<GraphRequestOptions, "path" | "method"> = {}): Promise<T> {
    return this.request<T>({ ...options, path, method: "GET" });
  }

  async post<T>(path: string, body?: unknown, options: Omit<GraphRequestOptions, "path" | "method" | "body"> = {}): Promise<T> {
    return this.request<T>({ ...options, path, method: "POST", body });
  }

  async patch<T>(path: string, body?: unknown, options: Omit<GraphRequestOptions, "path" | "method" | "body"> = {}): Promise<T> {
    return this.request<T>({ ...options, path, method: "PATCH", body });
  }

  async put<T>(path: string, body?: unknown, options: Omit<GraphRequestOptions, "path" | "method" | "body"> = {}): Promise<T> {
    return this.request<T>({ ...options, path, method: "PUT", body });
  }

  async delete(path: string, options: Omit<GraphRequestOptions, "path" | "method"> = {}): Promise<void> {
    await this.request<void>({ ...options, path, method: "DELETE" });
  }

  private buildUrl(options: GraphRequestOptions): string {
    const base = options.absoluteUrl ? options.path : joinUrl(this.options.baseUrl, options.path);
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

function combineSignals(first?: AbortSignal, second?: AbortSignal): AbortSignal | undefined {
  if (!first) return second;
  if (typeof AbortSignal !== "undefined" && "any" in AbortSignal) {
    return (AbortSignal as typeof AbortSignal & { any(signals: AbortSignal[]): AbortSignal }).any([first, second!]);
  }
  const controller = new AbortController();
  const abort = () => controller.abort();
  first.addEventListener("abort", abort, { once: true });
  second?.addEventListener("abort", abort, { once: true });
  return controller.signal;
}

function joinUrl(base: string, path: string): string {
  if (path.startsWith("http://") || path.startsWith("https://")) return path;
  return `${base.replace(/\/$/, "")}/${path.replace(/^\//, "")}`;
}

async function safeJson(response: Response): Promise<unknown> {
  try {
    return await response.json();
  } catch {
    return undefined;
  }
}

function parseRetryAfterOrBackoff(header: string | null, attempt: number): number {
  if (header) {
    const seconds = Number(header);
    if (!Number.isNaN(seconds)) return seconds * 1000;
    const date = Date.parse(header);
    if (!Number.isNaN(date)) return Math.max(250, date - Date.now());
  }
  return Math.min(8000, 500 * 2 ** attempt);
}
