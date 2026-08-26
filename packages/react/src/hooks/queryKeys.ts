/** Khóa React Query: tách cache theo site, tên thư viện, locale, folder, $orderby. */
export const queryKeys = {
  children: (
    siteId: string,
    library: string,
    locale: string,
    folderId: string,
    sortField = "",
    sortDirection = "",
  ) =>
    [
      "sp",
      siteId,
      library,
      locale,
      "children",
      folderId,
      sortField,
      sortDirection,
    ] as const,
  fields: (siteId: string, library: string, locale: string) =>
    ["sp", siteId, library, locale, "fields"] as const,
};
