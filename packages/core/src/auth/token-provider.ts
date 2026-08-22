/**
 * Tạo scope delegated SharePoint từ URL site.
 * Khớp quyền Entra **AllSites.Write** (đọc + ghi mọi site collection).
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aadhttpclient
 */
export function defaultSharePointScopes(siteUrl: string): string[] {
  const origin = new URL(siteUrl).origin;
  return [`${origin}/AllSites.Write`];
}

export const DEFAULT_SHAREPOINT_SCOPE_SUFFIX = "AllSites.Write";
