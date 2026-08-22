/**
 * Literal OData trong `'...'`: dấu nháy đơn phải nhân đôi.
 * Dùng cho Title list và ServerRelativeUrl — không encodeURIComponent cả chuỗi.
 */
export function escapeODataLiteral(value: string): string {
  return value.replace(/'/g, "''");
}
