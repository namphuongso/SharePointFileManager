import type { SharePointRestClient } from "../../rest/client";
import { SharePointError, SharePointErrorCode, isSharePointError } from "../../errors/sharepoint-error";
import type { CreateFolderOptions, LibraryContext, SharePointItem } from "../../types/models";
import type { RestFolder } from "../../types/rest";
import { assertValidItemName, requireUniqueId } from "../../utils";
import { encodeServerRelativePathArg } from "../../utils/odata-path-arg";
import { resolveFolderUrl } from "../folder/resolve-folder-url";
import { addFolderUsingPath } from "./add-folder-using-path";

/**
 * Tạo thư mục con (chỉ ghi). Parent = UniqueId folder hiện tại hoặc "root".
 * REST: resolve path → Folders/AddUsingPath (doc ResourcePath).
 */
export class CreateFolderService {
  constructor(
    private readonly rest: SharePointRestClient,
    private readonly getLibrary: () => Promise<LibraryContext>,
  ) {}

  async create(
    parentFolderId: string,
    name: string,
    options: CreateFolderOptions = {},
  ): Promise<SharePointItem> {
    assertValidItemName(name, "folder");
    const parentUrl = await resolveFolderUrl(
      this.rest,
      this.getLibrary,
      parentFolderId,
      options.signal,
    );
    const fullUrl = `${parentUrl.replace(/\/$/, "")}/${name}`;
    const folder = await this.rest.post<RestFolder>(addFolderUsingPath(fullUrl), {
      signal: options.signal,
    });
    if (folder.UniqueId && folder.Name) {
      return mapCreatedFolder(folder);
    }
    // Một số tenant trả body mỏng — GET lại theo path vừa tạo.
    return this.getByServerRelativeUrl(fullUrl, options.signal);
  }

  /**
   * Tạo folder theo ServerRelativeUrl đầy đủ — dùng khi upload thư mục (nhiều cấp).
   * Folder đã tồn tại (Conflict / already exists) → bỏ qua, không lỗi.
   */
  async ensureByServerRelativeUrl(
    serverRelativeUrl: string,
    options: CreateFolderOptions = {},
  ): Promise<void> {
    try {
      await this.rest.post(addFolderUsingPath(serverRelativeUrl), { signal: options.signal });
    } catch (error) {
      if (isSharePointError(error) && error.code === SharePointErrorCode.Conflict) return;
      throw error;
    }
  }

  /** Tạo chuỗi folder con dưới parentUniqueId theo đường dẫn tương đối `a/b/c`. */
  async ensureRelativePath(
    parentFolderId: string,
    relativeFolderPath: string,
    options: CreateFolderOptions = {},
  ): Promise<string> {
    const parentUrl = await resolveFolderUrl(
      this.rest,
      this.getLibrary,
      parentFolderId,
      options.signal,
    );
    const segments = relativeFolderPath.split("/").filter(Boolean);
    let current = parentUrl.replace(/\/$/, "");
    for (const segment of segments) {
      assertValidItemName(segment, "folder");
      current = `${current}/${segment}`;
      await this.ensureByServerRelativeUrl(current, options);
    }
    return current;
  }

  private async getByServerRelativeUrl(
    serverRelativeUrl: string,
    signal?: AbortSignal,
  ): Promise<SharePointItem> {
    const path = encodeServerRelativePathArg(serverRelativeUrl);
    const folder = await this.rest.get<RestFolder>(
      `web/GetFolderByServerRelativePath(decodedUrl='${path}')`,
      {
        query: { $select: "UniqueId,Name,ItemCount,TimeLastModified" },
        signal,
      },
    );
    return mapCreatedFolder(folder);
  }
}

function mapCreatedFolder(folder: RestFolder): SharePointItem {
  if (!folder.UniqueId || !folder.Name) {
    throw new SharePointError({
      code: SharePointErrorCode.Unknown,
      message: "Create folder response missing UniqueId or Name",
    });
  }
  return {
    id: requireUniqueId(folder.UniqueId, "folder"),
    name: folder.Name,
    type: "folder",
    lastModifiedDateTime: folder.TimeLastModified,
    childItemCount:
      folder.ItemCount !== undefined && folder.ItemCount !== null
        ? Number(folder.ItemCount)
        : 0,
  };
}
