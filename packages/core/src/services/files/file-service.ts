import type { SharePointRestClient } from "../../rest/client";
import type { LibraryContext } from "../../types/models";
import { resolveFileUrl } from "./resolve-file-url";

/** Tham số mở file trong tab mới. */
export interface OpenFileOptions {
  signal?: AbortSignal;
}

/**
 * Hành động mở file (chỉ đọc): UniqueId → URL ?web=1.
 * ?web=1 = preview SharePoint — Office Online, PDF, ảnh, txt; không preview được thì cho tải về.
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/general-development/working-with-folders-and-files-with-rest
 */
export class FileService {
  constructor(
    private readonly rest: SharePointRestClient,
    private readonly getLibrary: () => Promise<LibraryContext>,
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
}
