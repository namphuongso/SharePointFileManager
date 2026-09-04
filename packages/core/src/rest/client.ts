import { mapRestError } from "../errors/map-rest-error";
import { SharePointError, SharePointErrorCode } from "../errors/sharepoint-error";
import type { RestRequestOptions, SharePointRestClientOptions } from "../types/rest";
import {
  buildApiUrl,
  buildRestUrl,
  normalizeSiteUrl,
  parseSuccessBody,
  readErrorBody,
  sleep,
  throttleWaitMs,
} from "../utils";

/**
 * Client REST SharePoint — GET (duyệt / tải nhị phân) + POST (tạo folder / upload, Bearer OAuth).
 * OAuth: không cần X-RequestDigest. 401: token mới một lần. 429: chờ rồi thử ≤ 3 lần.
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service
 */
export class SharePointRestClient {
  constructor(private readonly options: SharePointRestClientOptions) {}

  get siteUrl(): string {
    return normalizeSiteUrl(this.options.siteUrl);
  }

  apiUrl(path: string): string {
    return buildApiUrl(this.siteUrl, path);
  }

  async get<T>(path: string, options: Omit<RestRequestOptions, "path" | "method"> = {}): Promise<T> {
    return this.request<T>({ ...options, path, method: "GET" });
  }

  /** GET URL tuyệt đối (trang sau `@odata.nextLink`). */
  async getUrl<T>(
    url: string,
    options: Omit<RestRequestOptions, "path" | "method" | "query" | "absoluteUrl"> = {},
  ): Promise<T> {
    return this.request<T>({ ...options, path: "", absoluteUrl: url, method: "GET" });
  }

  /**
   * GET nội dung nhị phân (vd. `GetFileById(...)/$value`).
   * Không dùng `get` — `parseSuccessBody` đọc text sẽ hỏng binary.
   */
  async getBlob(
    path: string,
    options: Omit<RestRequestOptions, "path" | "method" | "responseType"> = {},
  ): Promise<Blob> {
    return this.request<Blob>({ ...options, path, method: "GET", responseType: "blob" });
  }

  /** POST ghi (tạo folder, upload file, …). Body do caller truyền — không stringify trong client. */
  async post<T>(path: string, options: Omit<RestRequestOptions, "path" | "method"> = {}): Promise<T> {
    return this.request<T>({ ...options, path, method: "POST" });
  }

  async request<T>(options: RestRequestOptions): Promise<T> {
    const fetchImpl = this.options.fetchImpl ?? fetch;
    const method = options.method ?? "GET";
    const asBlob = options.responseType === "blob";
    let unauthorizedRetried = false;
    let throttleAttempt = 0;

    while (true) {
      throwIfCancelled(options.signal);

      const url = buildRestUrl((path) => this.apiUrl(path), options);
      const headers = new Headers(options.headers);
      if (!headers.has("Accept")) {
        headers.set(
          "Accept",
          asBlob ? "application/octet-stream" : "application/json;odata=nometadata",
        );
      }

      const token = await this.options.tokenProvider.getAccessToken({
        scopes: this.options.scopes,
        forceRefresh: unauthorizedRetried,
      });
      headers.set("Authorization", `Bearer ${token}`);

      let response: Response;
      try {
        response = await fetchImpl(url, {
          method,
          headers,
          body: options.body ?? undefined,
          signal: options.signal,
        });
      } catch (error) {
        throwIfCancelled(options.signal, error);
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

      const retryAfter = response.headers.get("Retry-After");
      if (response.status === 429 && throttleAttempt < 3) {
        throttleAttempt += 1;
        await sleep(throttleWaitMs(retryAfter, throttleAttempt));
        continue;
      }

      if (!response.ok) {
        throw mapRestError({
          status: response.status,
          body: await readErrorBody(response),
          retryAfter,
        });
      }

      if (asBlob) {
        return (await response.blob()) as T;
      }
      return parseSuccessBody<T>(response);
    }
  }
}

function throwIfCancelled(signal: AbortSignal | undefined, cause?: unknown): void {
  if (!signal?.aborted) return;
  throw new SharePointError({
    code: SharePointErrorCode.Cancelled,
    message: "Request was cancelled",
    cause,
  });
}

