import type { SharePointRestClient } from "../../rest/client";
import { SharePointError, SharePointErrorCode } from "../../errors/sharepoint-error";
import type { LibraryContext, SharePointItem, UploadFileOptions } from "../../types/models";
import type { RestFile } from "../../types/rest";
import { assertValidItemName, requireUniqueId } from "../../utils";
import { resolveFolderUrl } from "../folder/resolve-folder-url";
import { uploadFileByFolderPath } from "./upload-file-by-folder-path";

/**
 * Upload file vào folder (chỉ ghi). Parent = UniqueId hoặc "root".
 * REST chuẩn: UniqueId → ServerRelativeUrl → GetFolderByServerRelativePath/Files/add.
 */
export class UploadFileService {
  constructor(
    private readonly rest: SharePointRestClient,
    private readonly getLibrary: () => Promise<LibraryContext>,
  ) {}

  async upload(
    parentFolderId: string,
    content: Blob,
    options: UploadFileOptions,
  ): Promise<SharePointItem> {
    const folderUrl = await resolveFolderUrl(
      this.rest,
      this.getLibrary,
      parentFolderId,
      options.signal,
    );
    return this.uploadToFolderUrl(folderUrl, content, options);
  }

  /**
   * Upload vào folder đích theo ServerRelativeUrl (sau khi ensure cây thư mục).
   */
  async uploadToFolderUrl(
    folderServerRelativeUrl: string,
    content: Blob,
    options: UploadFileOptions,
  ): Promise<SharePointItem> {
    const fileName = options.fileName?.trim() || (content instanceof File ? content.name : "");
    assertValidItemName(fileName, "file");
    const overwrite = options.overwrite ?? false;

    const file = await this.rest.post<RestFile>(
      uploadFileByFolderPath(folderServerRelativeUrl, fileName, overwrite),
      {
        body: content,
        // Không ép MIME Office — octet-stream ổn định hơn với Files/add binary.
        headers: {
          "Content-Type": "application/octet-stream",
        },
        signal: options.signal,
      },
    );

    return mapUploadedFile(file, content.size);
  }

  /**
   * Upload nhiều file kèm webkitRelativePath (chọn thư mục trên máy).
   * Caller phải ensure folder trung gian trước (CreateFolderService.ensureRelativePath).
   */
  async uploadRelativeFile(
    parentFolderId: string,
    file: File,
    options: Omit<UploadFileOptions, "fileName"> = {},
  ): Promise<SharePointItem> {
    const relative = file.webkitRelativePath || file.name;
    const parts = relative.split("/").filter(Boolean);
    const leaf = parts.pop() ?? file.name;
    const parentUrl = await resolveFolderUrl(
      this.rest,
      this.getLibrary,
      parentFolderId,
      options.signal,
    );
    const folderUrl =
      parts.length === 0
        ? parentUrl.replace(/\/$/, "")
        : `${parentUrl.replace(/\/$/, "")}/${parts.join("/")}`;
    return this.uploadToFolderUrl(folderUrl, file, { ...options, fileName: leaf });
  }
}

function mapUploadedFile(file: RestFile, fallbackSize: number): SharePointItem {
  if (!file.UniqueId || !file.Name) {
    throw new SharePointError({
      code: SharePointErrorCode.Unknown,
      message: "Upload file response missing UniqueId or Name",
    });
  }

  const size =
    file.Length !== undefined && file.Length !== null ? Number(file.Length) : fallbackSize;

  return {
    id: requireUniqueId(file.UniqueId, "file"),
    name: file.Name,
    type: "file",
    size: Number.isFinite(size) ? size : fallbackSize,
    lastModifiedDateTime: file.TimeLastModified,
  };
}
