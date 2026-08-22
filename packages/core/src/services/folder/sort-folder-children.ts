import type { SharePointItem } from "../../types/models";

/** Folder trước file, rồi sort tên không phân biệt hoa thường. */
export function sortFolderChildren(items: SharePointItem[]): SharePointItem[] {
  return [...items].sort((a, b) => {
    if (a.type !== b.type) return a.type === "folder" ? -1 : 1;
    return a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
  });
}
