import type { SharePointRestClient } from "../../rest/client";
import { SharePointError, SharePointErrorCode } from "../../errors/sharepoint-error";
import type { LibraryContext } from "../../types/models";
import { normalizeGuid } from "../../utils/odata-path-arg";
import type { FolderService } from "../folder/folder";
import {
  buildFolderZip,
  type FolderZipPhase,
  type FolderZipProgress,
} from "./build-folder-zip";
import { fileDownloadPath } from "./file-download-path";
import { resolveFileUrl } from "./resolve-file-url";

export type { FolderZipPhase, FolderZipProgress };

/** Tham số mở / tải file hoặc folder (chỉ đọc). */
export interface OpenFileOptions {
  signal?: AbortSignal;
}

/** Tham số tải folder — kèm tên từ UI để đặt .zip. */
export interface DownloadFolderOptions extends OpenFileOptions {
  folderName?: string;
  onProgress?: (progress: FolderZipProgress) => void;
}

/** Kết quả tải nhị phân — UI gắn tên khi Save As. */
export interface DownloadFileResult {
  blob: Blob;
  /** Có khi tải folder (`.zip`); file dùng item.name từ UI. */
  fileName?: string;
}

/**
 * Hành động file/folder (chỉ đọc): mở preview ?web=1 hoặc tải /$value / zip folder.
 * Quyền OpenItems do caller check trước (PermissionService / useOpenItem / useDownloadItem).
 * Folder zip: listChildren đệ quy (cùng REST browse) + GetFileById/$value.
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest
 */
export class FileService {
  constructor(
    private readonly rest: SharePointRestClient,
    private readonly getLibrary: () => Promise<LibraryContext>,
    private readonly folders: FolderService,
  ) {}

  /** URL mở trong tab mới cho file theo UniqueId (không dùng cho folder). */
  async getOpenUrl(uniqueId: string, options?: OpenFileOptions): Promise<string> {
    const serverRelativeUrl = await resolveFileUrl(
      this.rest,
      this.getLibrary,
      uniqueId,
      options?.signal,
    );
    return `${this.rest.siteUrl}${serverRelativeUrl}?web=1`;
  }

  /**
   * Tải nội dung file theo UniqueId (`GetFileById(...)/$value`).
   * Không dùng cho folder — dùng `downloadFolder`.
   */
  async download(uniqueId: string, options?: OpenFileOptions): Promise<DownloadFileResult> {
    const library = await this.getLibrary();
    if (uniqueId === "root" || uniqueId === library.rootFolderUniqueId) {
      throw new SharePointError({
        code: SharePointErrorCode.NotFound,
        message: "Library root has no file content",
      });
    }
    const blob = await this.rest.getBlob(fileDownloadPath(normalizeGuid(uniqueId)), {
      signal: options?.signal,
    });
    return { blob };
  }

  /**
   * Tải cả thư mục thành `.zip` (listChildren đệ quy + GET /$value từng file).
   * Hỗ trợ alias `"root"` / rootFolderUniqueId.
   */
  async downloadFolder(
    uniqueId: string,
    options?: DownloadFolderOptions,
  ): Promise<DownloadFileResult> {
    const library = await this.getLibrary();
    const folderId =
      uniqueId === "root" ? library.rootFolderUniqueId : normalizeGuid(uniqueId);
    const result = await buildFolderZip(this.rest, this.folders, folderId, {
      signal: options?.signal,
      folderName: options?.folderName,
      onProgress: options?.onProgress,
    });
    return { blob: result.blob, fileName: result.fileName };
  }
}
