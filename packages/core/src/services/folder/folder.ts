import type { SharePointRestClient } from "../../rest/client";
import { mapRestFile, mapRestFolder } from "../../mappers/rest-item";
import type { LibraryContext, SharePointItem } from "../../types/models";
import type { ListChildrenOptions, RestFile, RestFolder } from "../../types/rest";
import { folderChildrenPath, resolveFolderUrl } from "./resolve-folder-url";
import { sortFolderChildren } from "./sort-folder-children";

/**
 * Liệt kê một cấp thư mục/file (chỉ đọc).
 * Path gốc lấy từ getLibrary (services/library).
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest
 */
export class FolderService {
  constructor(
    private readonly rest: SharePointRestClient,
    private readonly getLibrary: () => Promise<LibraryContext>,
  ) {}

  async listChildren(
    folderId: string,
    options: ListChildrenOptions = {},
  ): Promise<SharePointItem[]> {
    const url = await resolveFolderUrl(this.rest, this.getLibrary, folderId, options.signal);
    const base = folderChildrenPath(url);
    const top = options.top ?? 200;

    const [foldersResult, filesResult] = await Promise.all([
      this.rest.get<{ value?: RestFolder[] }>(`${base}/Folders`, {
        query: { $top: top },
        signal: options.signal,
      }),
      this.rest.get<{ value?: RestFile[] }>(`${base}/Files`, {
        query: { $top: top },
        signal: options.signal,
      }),
    ]);

    // Forms = folder hệ thống của document library, không hiện trên UI.
    const folders = (foldersResult.value ?? [])
      .filter((folder) => folder.Name && folder.Name !== "Forms")
      .map(mapRestFolder);
    const files = (filesResult.value ?? []).map(mapRestFile);

    return sortFolderChildren([...folders, ...files]);
  }
}
