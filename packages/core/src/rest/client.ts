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
 * Client REST SharePoint — chỉ GET (duyệt danh sách).
 * 401: lấy token mới một lần. 429: chờ rồi thử lại tối đa 3 lần.
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

  async request<T>(options: RestRequestOptions): Promise<T> {
    const fetchImpl = this.options.fetchImpl ?? fetch;
    const method = options.method ?? "GET";
    let unauthorizedRetried = false;
    let throttleAttempt = 0;

    while (true) {
      throwIfCancelled(options.signal);

      const url = buildRestUrl((path) => this.apiUrl(path), options);
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

