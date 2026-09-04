import type { SharePointRestClient } from "../../rest/client";
import { SharePointError, SharePointErrorCode } from "../../errors/sharepoint-error";
import type {
  CreateDocumentOptions,
  NewDocumentKind,
} from "../../types/create-document";
import type { LibraryContext, SharePointItem } from "../../types/models";
import type { RestFile } from "../../types/rest";
import { assertValidItemName, requireUniqueId } from "../../utils";
import { resolveFolderUrl } from "../folder/resolve-folder-url";
import { uploadFileByFolderPath } from "../upload/upload-file-by-folder-path";
import { blankDocumentContent, ensureDocumentFileName } from "./document-kinds";

/**
 * Tạo file trống (Office) bằng POST Files/add + template nhúng.
 * Path: UniqueId → ServerRelativeUrl → GetFolderByServerRelativePath/Files/add.
 */
export class CreateDocumentService {
  constructor(
    private readonly rest: SharePointRestClient,
    private readonly getLibrary: () => Promise<LibraryContext>,
  ) {}

  async create(
    parentFolderId: string,
    kind: NewDocumentKind,
    options: CreateDocumentOptions,
  ): Promise<SharePointItem> {
    const fileName = ensureDocumentFileName(kind, options.name);
    assertValidItemName(fileName, "file");
    const overwrite = options.overwrite ?? false;
    const folderUrl = await resolveFolderUrl(
      this.rest,
      this.getLibrary,
      parentFolderId,
      options.signal,
    );
    const { body } = blankDocumentContent(kind);

    const file = await this.rest.post<RestFile>(
      uploadFileByFolderPath(folderUrl, fileName, overwrite),
      {
        body,
        headers: { "Content-Type": "application/octet-stream" },
        signal: options.signal,
      },
    );

    if (!file.UniqueId || !file.Name) {
      throw new SharePointError({
        code: SharePointErrorCode.Unknown,
        message: "Create document response missing UniqueId or Name",
      });
    }

    const size =
      file.Length !== undefined && file.Length !== null ? Number(file.Length) : body.size;

    return {
      id: requireUniqueId(file.UniqueId, "file"),
      name: file.Name,
      type: "file",
      size: Number.isFinite(size) ? size : body.size,
      lastModifiedDateTime: file.TimeLastModified,
    };
  }
}
