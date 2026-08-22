import type { SharePointItem } from "../types/models";
import type { RestFile, RestFolder } from "../types/rest";
import { mapRestItem, requireUniqueId } from "../utils";

/** JSON SP.File → SharePointItem. */
export function mapRestFile(file: RestFile): SharePointItem {
  return mapRestItem(
    "file",
    file.Name ?? "file",
    requireUniqueId(file.UniqueId, "file"),
    file.TimeLastModified,
    file.Length,
  );
}

/** JSON SP.Folder → SharePointItem. */
export function mapRestFolder(folder: RestFolder): SharePointItem {
  return mapRestItem(
    "folder",
    folder.Name ?? "folder",
    requireUniqueId(folder.UniqueId, "folder"),
    folder.TimeLastModified,
  );
}
