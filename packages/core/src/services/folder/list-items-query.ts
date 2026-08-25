import { listItemExpand, listItemSelect } from "../fields/item-fields";
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
 * $select đúng cột đã có trên library (FieldService), không `*`.
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/use-odata-query-operations-in-sharepoint-rest-requests
 */
export function listItemsQuery(
  fileDirRef: string,
  top: number,
  fieldInternalNames: readonly string[],
  orderby: string,
) {
  const path = fileDirRef.replace(/\/+$/, "") || "/";
  return {
    $filter: `FileDirRef eq '${escapeODataLiteral(path)}'`,
    $top: top,
    $select: listItemSelect(fieldInternalNames),
    $expand: listItemExpand(fieldInternalNames),
    $orderby: orderby,
  };
}
