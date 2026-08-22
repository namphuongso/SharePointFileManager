import type { SharePointErrorCode } from "../types/errors";
import { SharePointErrorCode as ErrorCodes } from "./sharepoint-error";

/** HTTP status → mã lỗi thư viện. */
export function mapStatusToCode(status: number): SharePointErrorCode {
  if (status === 401) return ErrorCodes.Unauthorized;
  if (status === 403) return ErrorCodes.Forbidden;
  if (status === 404) return ErrorCodes.NotFound;
  if (status === 409 || status === 412) return ErrorCodes.Conflict;
  if (status === 413) return ErrorCodes.TooLarge;
  if (status === 429) return ErrorCodes.Throttled;
  if (status === 501) return ErrorCodes.Unsupported;
  return ErrorCodes.Unknown;
}
