import type { SharePointRestClient } from "../../rest/client";
import { mapRestListItem } from "../../mappers/rest-list-item";
import type { FolderChildrenPage, LibraryContext } from "../../types/models";
import type { ListChildrenOptions, RestListItem, RestODataCollection } from "../../types/rest";
import { parseODataCollection, resolveODataNextLink } from "../../utils";
import { DEFAULT_LIST_PAGE_SIZE, listItemsPath, listItemsQuery } from "./list-items-query";
import { resolveFolderUrl } from "./resolve-folder-url";

/**
 * Liệt kê một cấp file/folder qua list items (chỉ đọc).
 * Phân trang @odata.nextLink. Không $select — đủ cột trên mỗi dòng.
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/use-odata-query-operations-in-sharepoint-rest-requests
 */
export class FolderService {
  constructor(
    private readonly rest: SharePointRestClient,
    private readonly getLibrary: () => Promise<LibraryContext>,
  ) {}

  async listChildren(
    folderId: string,
    options: ListChildrenOptions = {},
  ): Promise<FolderChildrenPage> {
    const body = options.nextLink
      ? await this.rest.getUrl<RestODataCollection<RestListItem>>(options.nextLink, {
          signal: options.signal,
        })
      : await this.fetchFirstPage(folderId, options);

    const page = parseODataCollection(body);
    const items = page.value
      .map(mapRestListItem)
      .filter((item): item is NonNullable<typeof item> => item !== undefined);

    return {
      items,
      nextLink: resolveODataNextLink(page.nextLink, this.rest.siteUrl),
    };
  }

  private async fetchFirstPage(folderId: string, options: ListChildrenOptions) {
    const library = await this.getLibrary();
    const fileDirRef = await resolveFolderUrl(this.rest, this.getLibrary, folderId, options.signal);
    const top = options.top ?? DEFAULT_LIST_PAGE_SIZE;
    return this.rest.get<RestODataCollection<RestListItem>>(listItemsPath(library.listId), {
      query: listItemsQuery(fileDirRef, top),
      signal: options.signal,
    });
  }
}
