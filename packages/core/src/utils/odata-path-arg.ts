import { escapeODataLiteral } from "./odata";

/**
 * Giá trị leaf trong tham số OData quoted (`url='file name.docx'`, `add('Folder')`).
 * Escape dấu nháy rồi encode URI (khoảng trắng → %20).
 */
export function encodeODataQuotedValue(value: string): string {
  return encodeURIComponent(escapeODataLiteral(value));
}

/**
 * ServerRelativeUrl trong `decodedUrl='...'` / `decodedurl='...'`.
 * Giữ `/` giữa các segment (sample Microsoft); encode từng đoạn (%20, %, #, …).
 */
export function encodeServerRelativePathArg(path: string): string {
  return path
    .split("/")
    .map((segment) => encodeURIComponent(escapeODataLiteral(segment)))
    .join("/");
}

/** GUID trong GetFolderById / guid'...' — bỏ `{}`. */
export function normalizeGuid(id: string): string {
  return id.replace(/[{}]/g, "");
}
