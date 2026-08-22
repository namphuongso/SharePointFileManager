/**
 * Đưa path server-relative vào chuỗi OData: REST dùng `'path'`, dấu ' trong path phải thành `''`.
 * Không encodeURIComponent cả path — chỉ nhân đôi dấu nháy.
 */
export function encodeServerRelativeUrl(url: string): string {
  return url.replace(/'/g, "''");
}
