import type { SharePointItem, SharePointItemType } from "../types/models";
import type { RestFile, RestFolder } from "../types/rest";
import { requireUniqueId } from "../utils";

function mapRestItem(
  type: SharePointItemType,
  name: string,
  uniqueId: string,
  lastModifiedDateTime?: string,
  size?: number,
): SharePointItem {
  return { id: uniqueId, name, type, size, lastModifiedDateTime };
}

/** JSON SP.File → SharePointItem. Thiếu UniqueId thì không dùng làm key UI được. */
export function mapRestFile(file: RestFile): SharePointItem {
  return mapRestItem(
    "file",
    file.Name ?? "file",
    requireUniqueId(file.UniqueId, "file"),
    file.TimeLastModified,
    file.Length,
  );
}

/** JSON SP.Folder → SharePointItem. Thiếu UniqueId thì không dùng làm key UI được. */
export function mapRestFolder(folder: RestFolder): SharePointItem {
  return mapRestItem(
    "folder",
    folder.Name ?? "folder",
    requireUniqueId(folder.UniqueId, "folder"),
    folder.TimeLastModified,
  );
}
