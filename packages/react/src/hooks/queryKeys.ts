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
  childrenInfinite: (siteId: string, drive: string, folderId: string, expandListItem: boolean) =>
    ["sp", siteId, drive, "children-infinite", folderId, expandListItem] as const,
  listColumns: (siteId: string, drive: string) => ["sp", siteId, drive, "list-columns"] as const,
  accessibleLibraryItems: (siteId: string, drive: string) =>
    ["sp", siteId, drive, "accessible-library-items"] as const,
  listItemFields: (siteId: string, drive: string, itemId: string) =>
    ["sp", siteId, drive, "list-item-fields", itemId] as const,
  activities: (siteId: string, drive: string, itemId: string) =>
    ["sp", siteId, drive, "activities", itemId] as const,
  people: (query: string) => ["sp", "people", query] as const,
};
