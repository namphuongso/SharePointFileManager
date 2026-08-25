import { escapeODataLiteral } from "../../utils";

/** Path GET cột của view mặc định. GUID trong REST không dùng dấu ngoặc nhọn. */
export function viewFieldsPath(listId: string): string {
  const id = listId.replace(/[{}]/g, "");
  return `web/lists(guid'${escapeODataLiteral(id)}')/defaultView/viewfields`;
}
