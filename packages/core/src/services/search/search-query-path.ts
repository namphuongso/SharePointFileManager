/**
 * Absolute URL thư viện để dùng trong KQL Path: — siteUrl + ServerRelativeUrl gốc.
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/general-development/sharepoint-search-rest-api-overview
 */
export function libraryAbsoluteUrl(siteUrl: string, rootFolderServerRelativeUrl: string): string {
  const base = siteUrl.replace(/\/$/, "");
  const path = rootFolderServerRelativeUrl.startsWith("/")
    ? rootFolderServerRelativeUrl
    : `/${rootFolderServerRelativeUrl}`;
  return `${base}${path}`.replace(/\/+$/, "");
}

/** Path GET Search REST. */
export function searchQueryPath(): string {
  return "search/query";
}
