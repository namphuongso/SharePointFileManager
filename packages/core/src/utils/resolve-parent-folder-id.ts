import type { LibraryContext } from "../types/models";
import { normalizeGuid } from "./odata-path-arg";

/**
 * Alias "root" / UniqueId gốc thư viện → rootFolderUniqueId cho GetFolderById.
 * Folder con: UniqueId đã có trên breadcrumb.
 */
export async function resolveParentFolderUniqueId(
  getLibrary: () => Promise<LibraryContext>,
  folderId: string,
): Promise<string> {
  const library = await getLibrary();
  const normalized = normalizeGuid(folderId);
  if (folderId === "root" || normalized === normalizeGuid(library.rootFolderUniqueId)) {
    return normalizeGuid(library.rootFolderUniqueId);
  }
  return normalized;
}
