import type { SharePointRestClient } from "../../rest/client";
import {
  SharePointError,
  SharePointErrorCode,
} from "../../errors/sharepoint-error";
import type { LibraryContext } from "../../types/models";
import type { RestFolder } from "../../types/rest";

/**
 * UniqueId (breadcrumb) → ServerRelativeUrl cho $filter FileDirRef.
 * "root" hoặc UniqueId của library = path gốc, không gọi GetFolderById.
 */
export async function resolveFolderUrl(
  rest: SharePointRestClient,
  getLibrary: () => Promise<LibraryContext>,
  itemId: string,
  signal?: AbortSignal,
): Promise<string> {
  const library = await getLibrary();
  if (itemId === "root" || itemId === library.rootFolderUniqueId) {
    return library.rootFolderServerRelativeUrl;
  }

  const folder = await rest.get<RestFolder>(`web/GetFolderById('${itemId}')`, {
    query: { $select: "ServerRelativeUrl" },
    signal,
  });
  if (!folder.ServerRelativeUrl) {
    throw new SharePointError({
      code: SharePointErrorCode.NotFound,
      message: `Folder ${itemId} was not found`,
    });
  }
  return folder.ServerRelativeUrl;
}
