import type { ListChildrenSort } from "../../types/rest";

const SORTABLE_FIXED = new Set(["FileLeafRef", "Modified", "File_x0020_Size"]);
const PERSON_ORDERBY = new Set(["Author", "Editor"]);
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

/** Cột REST $orderby được: 3 cột cố định, Author/Editor, kiểu nguyên thủy. */
export function isSortableLibraryField(internalName: string, typeAsString?: string): boolean {
  if (SORTABLE_FIXED.has(internalName) || PERSON_ORDERBY.has(internalName)) return true;
  const type = typeAsString?.toLowerCase();
  if (!type || type === "computed") return false;
  return SORTABLE_TYPES.has(type);
}

/** Folder trước, rồi cột user hoặc FileLeafRef — giống All Documents. */
export function listItemsOrderby(
  sort: ListChildrenSort | undefined,
  typeAsString?: string,
): string {
  if (!sort) return "FSObjType desc,FileLeafRef";
  return `FSObjType desc,${orderbyProperty(sort.field, typeAsString)} ${sort.direction}`;
}

function orderbyProperty(field: string, typeAsString?: string): string {
  if (field === "ID") return "Id";
  if (PERSON_ORDERBY.has(field) || typeAsString?.toLowerCase() === "user") {
    return `${field}/Title`;
  }
  return field;
}
