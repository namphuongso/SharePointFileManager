/** Khóa React Query: tách cache theo site, tên thư viện, folder. */
export const queryKeys = {
  children: (siteId: string, library: string, folderId: string) =>
    ["sp", siteId, library, "children", folderId] as const,
};
