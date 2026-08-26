import { isSharePointError } from "./sharepoint-error";

/** SharePoint từ chối cột trong $select/$filter/$orderby. */
const MISSING_FIELD_PATTERN = /The field or property '([^']+)' does not exist/;
const MISSING_COLUMN_PATTERN = /(?:The )?column '([^']+)' does not exist/i;

/** Lấy tên cột bị SharePoint báo thiếu — dùng để loại cột "ghost" rồi thử lại. */
export function extractMissingField(error: unknown): string | undefined {
  if (!isSharePointError(error)) return undefined;
  return (
    MISSING_FIELD_PATTERN.exec(error.message)?.[1] ??
    MISSING_COLUMN_PATTERN.exec(error.message)?.[1]
  );
}
