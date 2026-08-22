import { escapeODataLiteral } from "../../utils";

/** Số dòng một trang — gần RowLimit view SharePoint. */
export const DEFAULT_LIST_PAGE_SIZE = 30;

/** Path GET items của list. GUID không dùng dấu ngoặc nhọn. */
export function listItemsPath(listId: string): string {
  const id = listId.replace(/[{}]/g, "");
  return `web/lists(guid'${escapeODataLiteral(id)}')/items`;
}

/**
 * Một cấp: FileDirRef = ServerRelativeUrl folder.
 * `$select=*` lấy đủ cột list; File/Folder/Author/Editor phải có trong `$select` khi `$expand`.
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/use-odata-query-operations-in-sharepoint-rest-requests
 */
export function listItemsQuery(fileDirRef: string, top: number) {
  const path = fileDirRef.replace(/\/+$/, "") || "/";
  return {
    $filter: `FileDirRef eq '${escapeODataLiteral(path)}'`,
    $top: top,
    $select:
      "*,File/UniqueId,File/Name,File/Length,File/TimeLastModified,Folder/UniqueId,Folder/Name,Folder/TimeLastModified,Author/Id,Author/Title,Editor/Id,Editor/Title",
    $expand: "File,Folder,Author,Editor",
    $orderby: "FSObjType desc,FileLeafRef",
  };
}
