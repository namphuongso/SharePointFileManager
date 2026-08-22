import { escapeODataLiteral } from "../../utils";

/** Path GET Fields của list. GUID trong REST không dùng dấu ngoặc nhọn. */
export function listFieldsPath(listId: string): string {
  const id = listId.replace(/[{}]/g, "");
  return `web/lists(guid'${escapeODataLiteral(id)}')/fields`;
}
