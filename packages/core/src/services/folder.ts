import type { SharePointRestClient } from "../rest/client";
import type { SharePointItem } from "../types/models";
import { SharePointError, SharePointErrorCode } from "../errors/sharepoint-error";
import { mapRestFile, mapRestFolder, type RestFile, type RestFolder } from "../mappers/rest-item";
import { encodeServerRelativeUrl, type LibraryContext } from "./library";

/**
 * Folder/file listing via SharePoint REST (read-only).
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest
 */
export class FolderService {
  constructor(
    private readonly rest: SharePointRestClient,
    private readonly getLibrary: () => Promise<LibraryContext>,
  ) {}

  async listChildren(
    folderId: string,
    options: { top?: number; signal?: AbortSignal } = {},
  ): Promise<SharePointItem[]> {
    const url = await this.resolveFolderUrl(folderId, options.signal);
    const encoded = encodeServerRelativeUrl(url);
    const top = options.top ?? 200;

    const [foldersResult, filesResult] = await Promise.all([
      this.rest.get<{ value?: RestFolder[] }>(`web/GetFolderByServerRelativeUrl('${encoded}')/Folders`, {
        query: {
          $select: "UniqueId,Name,TimeLastModified",
          $top: top,
        },
        signal: options.signal,
      }),
      this.rest.get<{ value?: RestFile[] }>(`web/GetFolderByServerRelativeUrl('${encoded}')/Files`, {
        query: {
          $select: "UniqueId,Name,Length,TimeLastModified",
          $top: top,
        },
        signal: options.signal,
      }),
    ]);

    const folders = (foldersResult.value ?? [])
      .filter((folder) => folder.Name && folder.Name !== "Forms")
      .map(mapRestFolder);
    const files = (filesResult.value ?? []).map(mapRestFile);

    return [...folders, ...files].sort((a, b) => {
      if (a.type !== b.type) return a.type === "folder" ? -1 : 1;
      return a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
    });
  }

  private async resolveFolderUrl(itemId: string, signal?: AbortSignal): Promise<string> {
    const library = await this.getLibrary();
    if (itemId === "root" || itemId === library.rootFolderUniqueId) {
      return library.rootFolderServerRelativeUrl;
    }

    const folder = await this.rest.get<RestFolder>(`web/GetFolderById('${itemId}')`, {
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
}
