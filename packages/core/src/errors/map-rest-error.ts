import { SharePointError, SharePointErrorCode } from "./sharepoint-error";

interface RestErrorBody {
  error?: {
    code?: string;
    message?: { value?: string } | string;
  };
  "odata.error"?: {
    code?: string;
    message?: { value?: string };
  };
}

const REST_CODE_MAP: Record<string, SharePointErrorCode> = {
  "-2147024891, System.UnauthorizedAccessException": SharePointErrorCode.Forbidden,
  accessDenied: SharePointErrorCode.Forbidden,
  unauthorized: SharePointErrorCode.Unauthorized,
  "-2147024894, System.IO.FileNotFoundException": SharePointErrorCode.NotFound,
  itemNotFound: SharePointErrorCode.NotFound,
  "-2130575257, Microsoft.SharePoint.SPException": SharePointErrorCode.Conflict,
};

export function mapStatusToCode(status: number): SharePointErrorCode {
  if (status === 401) return SharePointErrorCode.Unauthorized;
  if (status === 403) return SharePointErrorCode.Forbidden;
  if (status === 404) return SharePointErrorCode.NotFound;
  if (status === 409 || status === 412) return SharePointErrorCode.Conflict;
  if (status === 413) return SharePointErrorCode.TooLarge;
  if (status === 429) return SharePointErrorCode.Throttled;
  if (status === 501) return SharePointErrorCode.Unsupported;
  return SharePointErrorCode.Unknown;
}

export function parseRetryAfterMs(header: string | null): number | undefined {
  if (!header) return undefined;
  const seconds = Number(header);
  if (!Number.isNaN(seconds)) return seconds * 1000;
  const date = Date.parse(header);
  if (!Number.isNaN(date)) return Math.max(0, date - Date.now());
  return undefined;
}

export function mapRestError(input: {
  status: number;
  body?: unknown;
  retryAfter?: string | null;
  fallbackMessage?: string;
}): SharePointError {
  const body = input.body as RestErrorBody | undefined;
  const err = body?.error ?? body?.["odata.error"];
  const rawMessage = err?.message;
  const message =
    (typeof rawMessage === "string" ? rawMessage : rawMessage?.value) ??
    input.fallbackMessage ??
    `SharePoint REST request failed with status ${input.status}`;
  const restCode = err?.code;
  const code = (restCode && REST_CODE_MAP[restCode]) || mapStatusToCode(input.status);

  return new SharePointError({
    code,
    message,
    status: input.status,
    restCode,
    retryAfterMs: parseRetryAfterMs(input.retryAfter ?? null),
  });
}
