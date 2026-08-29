import { escapeODataLiteral } from "../../utils";

/** GUID OData — bỏ `{}` nếu REST trả kèm. */
function normalizeGuid(id: string): string {
  return id.replace(/[{}]/g, "");
}

/** Library — user hiện tại; list không có UniqueId trong model app. */
export function libraryPermissionsPath(listId: string): string {
  return `web/lists(guid'${escapeODataLiteral(normalizeGuid(listId))}')/effectiveBasePermissions`;
}

/**
 * List item (file/folder trong document library) theo UniqueId.
 * GetItemByUniqueId [Remote] — khớp File.UniqueId / Folder.UniqueId từ list items.
 * @see https://learn.microsoft.com/en-us/dotnet/api/microsoft.sharepoint.client.list.getitembyuniqueid
 */
export function listItemPermissionsPath(listId: string, uniqueId: string): string {
  const list = escapeODataLiteral(normalizeGuid(listId));
  const item = escapeODataLiteral(normalizeGuid(uniqueId));
  return `web/lists(guid'${list}')/GetItemByUniqueId(guid'${item}')/effectiveBasePermissions`;
}

/**
 * Folder theo UniqueId — root không phải list item nên dùng quyền library.
 * Alias "root" hoặc rootFolderUniqueId → effectiveBasePermissions trên list.
 */
export function folderPermissionsPath(
  listId: string,
  uniqueId: string,
  rootFolderUniqueId: string,
): string {
  const normalizedRoot = normalizeGuid(rootFolderUniqueId);
  if (uniqueId === "root" || normalizeGuid(uniqueId) === normalizedRoot) {
    return libraryPermissionsPath(listId);
  }
  return listItemPermissionsPath(listId, uniqueId);
}
