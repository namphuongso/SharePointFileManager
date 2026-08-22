import type { SharePointErrorCode as ErrorCode, SharePointErrorOptions } from "../types/errors";

export const SharePointErrorCode = {
  Unauthorized: "Unauthorized",
  Forbidden: "Forbidden",
  NotFound: "NotFound",
  Conflict: "Conflict",
  TooLarge: "TooLarge",
  Throttled: "Throttled",
  NetworkError: "NetworkError",
  Unsupported: "Unsupported",
  Cancelled: "Cancelled",
  Unknown: "Unknown",
} as const satisfies Record<ErrorCode, ErrorCode>;

/** Lỗi SharePoint REST: mã ổn định + HTTP status + Retry-After khi bị 429. */

export class SharePointError extends Error {
  readonly code: ErrorCode;
  readonly status?: number;
  readonly restCode?: string;
  readonly retryAfterMs?: number;

  constructor(options: SharePointErrorOptions) {
    super(options.message, { cause: options.cause });
    this.name = "SharePointError";
    this.code = options.code;
    this.status = options.status;
    this.restCode = options.restCode;
    this.retryAfterMs = options.retryAfterMs;
  }
}

export function isSharePointError(error: unknown): error is SharePointError {
  return error instanceof SharePointError;
}
