import type { SharePointItem } from "../types/models";

/**
 * Folder view over security-trimmed list items — same idea as SharePoint library UI.
 * Items the user can access stay under their parent folder when that folder is also visible.
 */
export function itemsVisibleInFolder(items: SharePointItem[], folderId: string): SharePointItem[] {
  const inFolder = items.filter((item) => item.parentId === folderId);
  if (inFolder.length > 0) return inFolder;
  if (items.some((item) => item.id === folderId)) {
    return items.filter((item) => item.parentId === folderId);
  }

  const folders = items.filter((item) => item.type === "folder");
  const folderIds = new Set(folders.map((folder) => folder.id));
  const files = items.filter((item) => item.type === "file");
  const filesWithoutVisibleParent = files.filter(
    (file) => !file.parentId || !folderIds.has(file.parentId),
  );
  return [...folders, ...filesWithoutVisibleParent];
}
