export type SharePointErrorCode =
  | "Unauthorized"
  | "Forbidden"
  | "NotFound"
  | "Conflict"
  | "TooLarge"
  | "Throttled"
  | "NetworkError"
  | "Unsupported"
  | "Cancelled"
  | "Unknown";

export interface SharePointErrorOptions {
  code: SharePointErrorCode;
  message: string;
  status?: number;
  restCode?: string;
  retryAfterMs?: number;
  cause?: unknown;
}
