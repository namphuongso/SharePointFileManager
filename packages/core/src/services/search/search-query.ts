import { libraryAbsoluteUrl } from "./search-query-path";
import { searchManagedProperty, searchSortlist } from "./search-sort";
import type { ListChildrenSort } from "../../types/rest";

/** Số dòng một trang Search — gần browse list. */
export const DEFAULT_SEARCH_PAGE_SIZE = 30;

/** Managed properties tối thiểu để map SharePointItem. */
export const SEARCH_BASE_PROPERTIES = [
  "Title",
  "Filename",
  "Path",
  "UniqueId",
  "IsContainer",
  "Size",
  "LastModifiedTime",
  "FileExtension",
  "Author",
  "ModifiedBy",
  "Created",
  "FolderChildCount",
] as const;

export interface LibrarySearchQueryInput {
  siteUrl: string;
  rootFolderServerRelativeUrl: string;
  rowLimit: number;
  startRow: number;
  /** Cột option đang hiện — InternalName list → managed property. */
  fieldInternalNames?: readonly string[];
  sort?: ListChildrenSort;
}

/**
 * Query params GET /_api/search/query — giới hạn Path thư viện; sortlist / selectproperties.
 * Search REST GET bắt buộc bọc chuỗi trong dấu nháy đơn: querytext='...'.
 */
export function librarySearchQuery(input: LibrarySearchQueryInput): Record<string, string | number> {
  const libraryUrl = libraryAbsoluteUrl(input.siteUrl, input.rootFolderServerRelativeUrl);
  const formsPath = `${libraryUrl}/Forms`;
  const kql = `Path:"${libraryUrl}*" -Path:"${formsPath}*"`;

  const selected = new Set<string>(SEARCH_BASE_PROPERTIES);
  for (const name of input.fieldInternalNames ?? []) {
    selected.add(searchManagedProperty(name));
  }
  const selectproperties = [...selected].join(",");
  const sortlist = searchSortlist(input.sort);

  const query: Record<string, string | number> = {
    querytext: `'${kql}'`,
    selectproperties: `'${selectproperties}'`,
    rowlimit: input.rowLimit,
    startrow: input.startRow,
    trimduplicates: "false",
  };
  if (sortlist) {
    query.sortlist = `'${sortlist}'`;
  }
  return query;
}
