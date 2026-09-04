import type { SharePointRestClient } from "../../rest/client";
import { SharePointError, SharePointErrorCode } from "../../errors/sharepoint-error";
import type {
  DeleteItemOptions,
  LibraryContext,
  SharePointItemType,
} from "../../types/models";
import { normalizeGuid } from "../../utils/odata-path-arg";
import { listItemRecyclePath } from "./delete-item-path";

/**
 * Xóa mềm file/folder (Recycle Bin) theo UniqueId.
 * Quyền DeleteListItems do caller check trước (PermissionService / menu).
 */
export class DeleteItemService {
  constructor(
    private readonly rest: SharePointRestClient,
    private readonly getLibrary: () => Promise<LibraryContext>,
  ) {}

  /**
   * POST recycle — không xóa cứng (HTTP DELETE).
   * Cấm thư viện gốc (`"root"` / rootFolderUniqueId).
   */
  async delete(
    type: SharePointItemType,
    uniqueId: string,
    options: DeleteItemOptions = {},
  ): Promise<void> {
    if (type !== "file" && type !== "folder") {
      throw new SharePointError({
        code: SharePointErrorCode.Unsupported,
        message: "Only file or folder can be deleted",
      });
    }

    const library = await this.getLibrary();
    if (uniqueId === "root" || normalizeGuid(uniqueId) === normalizeGuid(library.rootFolderUniqueId)) {
      throw new SharePointError({
        code: SharePointErrorCode.Unsupported,
        message: "Cannot delete library root folder",
      });
    }

    await this.rest.post(listItemRecyclePath(library.listId, uniqueId), {
      signal: options.signal,
    });
  }
}
