import type { ListChildrenSort } from "../../types/rest";

/** Cột cố định UI — $orderby trực tiếp trên list items được. */
const SORTABLE_FIXED = new Set(["FileLeafRef", "Modified"]);
const PERSON_ORDERBY = new Set(["Author", "Editor"]);

/**
 * Tên hệ thống thư viện: REST $orderby hay fail (computed / navigation / virtual),
 * kể cả khi TypeAsString UI gắn nhầm (vd. Size → Number).
 */
const UNSORTABLE_INTERNAL = new Set([
  "File_x0020_Size",
  "FileSize",
  "FileSizeDisplay",
  "DocIcon",
  "LinkFilename",
  "LinkFilenameNoMenu",
  "ContentType",
  "ContentTypeId",
  "ItemChildCount",
  "FolderChildCount",
  "Edit",
  "HTML_x0020_File_x0020_Type",
  "EncodedAbsUrl",
  "BaseName",
  "FileRef",
  "FileDirRef",
  "UniqueId",
  "FSObjType",
  "GUID",
]);

const SORTABLE_TYPES = new Set([
  "text",
  "number",
  "integer",
  "counter",
  "currency",
  "datetime",
  "boolean",
  "choice",
  "guid",
]);

/** Cột REST $orderby được: tên/ngày sửa, Author/Editor (khi còn trên list), kiểu nguyên thủy. */
export function isSortableLibraryField(
  internalName: string,
  typeAsString?: string,
): boolean {
  if (UNSORTABLE_INTERNAL.has(internalName)) return false;
  if (SORTABLE_FIXED.has(internalName) || PERSON_ORDERBY.has(internalName))
    return true;
  const type = typeAsString?.toLowerCase();
  if (
    !type ||
    type === "computed" ||
    type === "calculated" ||
    type === "note" ||
    type === "lookup" ||
    type === "lookupmulti" ||
    type === "usermulti" ||
    type === "multichoice" ||
    type === "url" ||
    type === "taxonomyfieldtype" ||
    type === "taxonomyfieldtypemulti" ||
    type === "attachments" ||
    type === "file" ||
    type === "user"
  ) {
    // User thường: chỉ Author/Editor (PERSON_ORDERBY). User khác cần /Title + $expand riêng.
    return false;
  }
  return SORTABLE_TYPES.has(type);
}

/** Search sortlist: Name/Modified/Size + Author/Editor + kiểu nguyên thủy (giống browse). */
export function isSortableSearchField(
  internalName: string,
  typeAsString?: string,
): boolean {
  if (internalName === "File_x0020_Size") return true;
  return isSortableLibraryField(internalName, typeAsString);
}

/**
 * Folder trước, rồi cột user hoặc FileLeafRef — giống All Documents.
 * `typeAsString` phải từ field còn trên list (không computed/ghost); thiếu thì không sort person/custom.
 */
export function listItemsOrderby(
  sort: ListChildrenSort | undefined,
  typeAsString?: string,
): string {
  if (!sort) return "FSObjType desc,FileLeafRef";
  if (SORTABLE_FIXED.has(sort.field)) {
    return `FSObjType desc,${orderbyProperty(sort.field, typeAsString)} ${sort.direction}`;
  }
  // Author/Editor và cột custom: cần metadata field còn sống — tránh $orderby=Author/Title khi đã exclude.
  if (!typeAsString || !isSortableLibraryField(sort.field, typeAsString)) {
    return "FSObjType desc,FileLeafRef";
  }
  return `FSObjType desc,${orderbyProperty(sort.field, typeAsString)} ${sort.direction}`;
}

function orderbyProperty(field: string, typeAsString?: string): string {
  if (field === "ID") return "Id";
  if (PERSON_ORDERBY.has(field) || typeAsString?.toLowerCase() === "user") {
    return `${field}/Title`;
  }
  return field;
}
