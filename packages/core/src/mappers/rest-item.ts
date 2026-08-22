import type { SharePointItem } from "../types/models";
import type { RestFile, RestFolder } from "../types/rest";

/** JSON SP.File → SharePointItem. Thiếu UniqueId thì không dùng làm key UI được. */
export function mapRestFile(file: RestFile): SharePointItem {
  const name = file.Name ?? "file";
  const id = file.UniqueId;
  if (!id) throw new Error("SharePoint file is missing UniqueId");
  return {
    id,
    name,
    type: "file",
    size: file.Length,
    lastModifiedDateTime: file.TimeLastModified,
  };
}

/** JSON SP.Folder → SharePointItem. */
export function mapRestFolder(folder: RestFolder): SharePointItem {
  const name = folder.Name ?? "folder";
  const id = folder.UniqueId;
  if (!id) throw new Error("SharePoint folder is missing UniqueId");
  return {
    id,
    name,
    type: "folder",
    lastModifiedDateTime: folder.TimeLastModified,
  };
}
