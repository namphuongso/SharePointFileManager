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
} as const;

export type SharePointErrorCode =
  (typeof SharePointErrorCode)[keyof typeof SharePointErrorCode];

export class SharePointError extends Error {
  readonly code: SharePointErrorCode;
  readonly status?: number;
  readonly restCode?: string;
  readonly retryAfterMs?: number;

  constructor(options: {
    code: SharePointErrorCode;
    message: string;
    status?: number;
    restCode?: string;
    retryAfterMs?: number;
    cause?: unknown;
  }) {
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
