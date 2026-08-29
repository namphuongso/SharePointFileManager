import type { ListChildrenSort } from "../../types/rest";

/**
 * InternalName cột UI → managed property Search (sortlist).
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/general-development/sharepoint-search-rest-api-overview
 */
const SORT_MANAGED: Record<string, string> = {
  FileLeafRef: "Filename",
  Modified: "LastModifiedTime",
  File_x0020_Size: "Size",
  Author: "Author",
  Editor: "ModifiedBy",
  Created: "Created",
  Title: "Title",
};

/** sortlist GET Search — `'Property:ascending'` hoặc undefined (rank mặc định). */
export function searchSortlist(sort: ListChildrenSort | undefined): string | undefined {
  if (!sort) return undefined;
  const property = SORT_MANAGED[sort.field] ?? sort.field;
  const direction = sort.direction === "asc" ? "ascending" : "descending";
  return `${property}:${direction}`;
}

/**
 * InternalName cột option → managed property lấy trong selectproperties.
 * Author/Editor có tên crawl khác list.
 */
export function searchManagedProperty(internalName: string): string {
  if (internalName === "Editor") return "ModifiedBy";
  if (internalName === "Author") return "Author";
  if (internalName === "Created") return "Created";
  if (internalName === "Modified") return "LastModifiedTime";
  if (internalName === "FileLeafRef") return "Filename";
  if (internalName === "File_x0020_Size") return "Size";
  if (internalName === "ItemChildCount") return "FolderChildCount";
  return internalName;
}
