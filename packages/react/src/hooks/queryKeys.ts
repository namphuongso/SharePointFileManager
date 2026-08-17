export const queryKeys = {
  site: (siteUrl: string) => ["sp", "site", siteUrl] as const,
  lists: (siteId: string) => ["sp", siteId, "lists"] as const,
  drives: (siteId: string) => ["sp", siteId, "drives"] as const,
  children: (siteId: string, drive: string, folderId: string) =>
    ["sp", siteId, drive, "children", folderId] as const,
  item: (siteId: string, drive: string, itemId: string) => ["sp", siteId, drive, "item", itemId] as const,
  permissions: (siteId: string, drive: string, itemId: string) =>
    ["sp", siteId, drive, "permissions", itemId] as const,
  versions: (siteId: string, drive: string, itemId: string) =>
    ["sp", siteId, drive, "versions", itemId] as const,
  search: (siteId: string, drive: string, query: string) => ["sp", siteId, drive, "search", query] as const,
};
