/** Khóa React Query: tách cache theo site, tên thư viện, locale, folder. */
export const queryKeys = {
  children: (siteId: string, library: string, locale: string, folderId: string) =>
    ["sp", siteId, library, locale, "children", folderId] as const,
  fields: (siteId: string, library: string, locale: string) =>
    ["sp", siteId, library, locale, "fields"] as const,
};
