import type { SharePointRestClient } from "../rest/client";
import { SharePointError, SharePointErrorCode } from "../errors/sharepoint-error";

export interface LibraryContext {
  listId: string;
  listTitle: string;
  rootFolderServerRelativeUrl: string;
  rootFolderUniqueId: string;
  entityTypeName?: string;
}

interface RestList {
  Id?: string;
  Title?: string;
  EntityTypeName?: string;
  RootFolder?: { ServerRelativeUrl?: string; UniqueId?: string; Name?: string };
}

const LIST_SELECT =
  "Id,Title,EntityTypeName,RootFolder/ServerRelativeUrl,RootFolder/UniqueId,RootFolder/Name";

function toLibraryContext(list: RestList, fallbackTitle?: string): LibraryContext {
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
    listTitle: list.Title ?? fallbackTitle ?? "Documents",
    rootFolderServerRelativeUrl: rootUrl,
    rootFolderUniqueId: rootId,
    entityTypeName: list.EntityTypeName,
  };
}

async function findDocumentLibraryByName(
  rest: SharePointRestClient,
  libraryName: string,
): Promise<RestList | undefined> {
  const needle = libraryName.trim().toLowerCase();
  const result = await rest.get<{ value?: RestList[] }>("web/lists", {
    query: {
      $filter: "BaseTemplate eq 101",
      $top: 200,
      $expand: "RootFolder",
      $select: LIST_SELECT,
    },
  });

  return (result.value ?? []).find((list) => {
    const title = list.Title?.trim().toLowerCase();
    const folderName = list.RootFolder?.Name?.trim().toLowerCase();
    const entity = list.EntityTypeName?.trim().toLowerCase();
    return title === needle || folderName === needle || entity === needle;
  });
}

/**
 * Resolve document library via SharePoint REST lists API.
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
 */
export async function resolveLibrary(
  rest: SharePointRestClient,
  options: { libraryName?: string; listId?: string },
): Promise<LibraryContext> {
  let list: RestList;
  if (options.listId) {
    list = await rest.get<RestList>(`web/lists(guid'${options.listId}')`, {
      query: { $expand: "RootFolder", $select: LIST_SELECT },
    });
  } else if (options.libraryName) {
    const title = options.libraryName.replace(/'/g, "''");
    try {
      list = await rest.get<RestList>(`web/lists/getbytitle('${title}')`, {
        query: { $expand: "RootFolder", $select: LIST_SELECT },
      });
    } catch {
      // Title có thể khác URL folder name (vd. đổi tên hiển thị) — tìm theo Title / RootFolder.Name.
      const matched = await findDocumentLibraryByName(rest, options.libraryName);
      if (!matched) {
        throw new SharePointError({
          code: SharePointErrorCode.NotFound,
          message: `Không tìm thấy thư viện "${options.libraryName}" trên site ${rest.siteUrl}. Kiểm tra VITE_SPFM_SITE_URL và VITE_SPFM_LIBRARY_NAME.`,
        });
      }
      list = matched;
    }
  } else {
    // Default document library
    list = await rest
      .get<RestList>("web/lists/getbytitle('Documents')", {
        query: { $expand: "RootFolder", $select: LIST_SELECT },
      })
      .catch(async () => {
        const result = await rest.get<{ value?: RestList[] }>("web/lists", {
          query: {
            $filter: "BaseTemplate eq 101",
            $top: 1,
            $expand: "RootFolder",
            $select: LIST_SELECT,
          },
        });
        const first = result.value?.[0];
        if (!first) {
          throw new SharePointError({
            code: SharePointErrorCode.NotFound,
            message: "No document library was found on this site",
          });
        }
        return first;
      });
  }

  return toLibraryContext(list, options.libraryName);
}

export function encodeServerRelativeUrl(url: string): string {
  return url.replace(/'/g, "''");
}
