import { SharePointError, SharePointErrorCode } from "../../errors/sharepoint-error";
import type { LibraryContext } from "../../types/models";
import type { RestList } from "../../types/rest";

/** REST list + RootFolder → context cho FolderService. Thiếu Id/path/UniqueId thì coi như không resolve được. */
export function toLibraryContext(list: RestList, fallbackTitle: string): LibraryContext {
  const rootUrl = list.RootFolder?.ServerRelativeUrl;
  const rootId = list.RootFolder?.UniqueId;
  if (!list.Id || !rootUrl || !rootId) {
    throw new SharePointError({
      code: SharePointErrorCode.NotFound,
      message: "SharePoint library root folder could not be resolved",
    });
  }

  return {
    listId: list.Id,
    listTitle: list.Title ?? fallbackTitle,
    rootFolderServerRelativeUrl: rootUrl,
    rootFolderUniqueId: rootId,
    entityTypeName: list.EntityTypeName,
  };
}
