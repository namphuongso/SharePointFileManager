/** Cột cố định của bảng UI; tên hiển thị vẫn lấy từ SharePoint theo locale. */
export const FIXED_LIBRARY_FIELD_NAMES = new Set(["FileLeafRef", "Modified", "File_x0020_Size"]);

/** Field schema/virtual: có metadata nhưng không select trực tiếp qua /items. */
const NON_SELECTABLE_ITEM_FIELD_PREFIXES = ["_"];
const NON_SELECTABLE_ITEM_FIELD_NAMES = new Set([
  "ComplianceAssetId",
  "FolderChildCount",
  "DocIcon",
  "ItemChildCount",
  "FileLeafRef",
  "FileSize",
]);

/**
 * Ba cột trên đã render từ File/Folder expand, nên không đưa vào $select items.
 * Field bắt đầu bằng `_` thường là schema/virtual và cũng không select được.
 */
export function selectableItemFieldNames(
  internalNames: readonly string[],
): readonly string[] {
  return internalNames.filter((name) => {
    if (FIXED_LIBRARY_FIELD_NAMES.has(name) || NON_SELECTABLE_ITEM_FIELD_NAMES.has(name)) {
      return false;
    }
    return !NON_SELECTABLE_ITEM_FIELD_PREFIXES.some((prefix) => name.startsWith(prefix));
  });
}

const LOOKUP_SELECT: Record<string, string> = {
  Author: "Author/Id,Author/Title",
  Editor: "Editor/Id,Editor/Title",
};

/** $expand File/Folder luôn; Author/Editor chỉ khi có trong $select (SharePoint bắt buộc cặp này). */
export function listItemExpand(internalNames: readonly string[]): string {
  const lookups = internalNames.filter((name) => name in LOOKUP_SELECT);
  return ["File", "Folder", ...lookups].join(",");
}

/** $select items theo cột trả về từ SharePoint + identity File/Folder. */
export function listItemSelect(internalNames: readonly string[]): string {
  const selectedFields = internalNames.map((name) => {
    const lookup = LOOKUP_SELECT[name];
    if (lookup) return lookup;
    return name === "ID" ? "Id" : name;
  });

  return [
    "GUID",
    "FSObjType",
    "FileDirRef",
    ...selectedFields,
    "File/UniqueId,File/Name,File/Length,File/TimeLastModified",
    "Folder/UniqueId,Folder/Name,Folder/ItemCount,Folder/TimeLastModified",
  ].join(",");
}
