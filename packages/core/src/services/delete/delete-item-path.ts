import { escapeODataLiteral } from "../../utils";
import { normalizeGuid } from "../../utils/odata-path-arg";

/**
 * Soft-delete list item (file/folder) theo UniqueId → Recycle Bin.
 * Cùng GetItemByUniqueId như quyền; ListItem.recycle (Microsoft).
 * @see https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-visio/jj247053(v=office.15)
 */
export function listItemRecyclePath(listId: string, uniqueId: string): string {
  const list = escapeODataLiteral(normalizeGuid(listId));
  const item = escapeODataLiteral(normalizeGuid(uniqueId));
  return `web/lists(guid'${list}')/GetItemByUniqueId(guid'${item}')/recycle()`;
}
