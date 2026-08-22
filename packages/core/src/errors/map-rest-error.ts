import type { MapRestErrorInput, RestErrorBody } from "../types/rest";
import type { SharePointErrorCode } from "../types/errors";
import { SharePointError, SharePointErrorCode as ErrorCodes } from "./sharepoint-error";
import { mapStatusToCode } from "./map-status-to-code";
import { parseRetryAfterMs } from "../utils";

/** Một số mã OData SharePoint không đi kèm HTTP status đúng ý nghĩa. */
const REST_CODE_MAP: Record<string, SharePointErrorCode> = {
  "-2147024891, System.UnauthorizedAccessException": ErrorCodes.Forbidden,
  accessDenied: ErrorCodes.Forbidden,
  unauthorized: ErrorCodes.Unauthorized,
  "-2147024894, System.IO.FileNotFoundException": ErrorCodes.NotFound,
  itemNotFound: ErrorCodes.NotFound,
  "-2130575257, Microsoft.SharePoint.SPException": ErrorCodes.Conflict,
};

/** JSON lỗi REST → SharePointError (mã, message, Retry-After). */
export function mapRestError(input: MapRestErrorInput): SharePointError {
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
