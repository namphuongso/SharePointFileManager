/** Bỏ khoảng trắng đầu/cuối và slash cuối URL site. */
export function normalizeSiteUrl(url: string): string {
  return url.trim().replace(/\/$/, "");
}
