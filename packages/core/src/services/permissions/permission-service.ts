import type { SharePointRestClient } from "../../rest/client";
import type { LibraryContext } from "../../types/models";
import type { EffectiveBasePermissionsDto, ItemCapabilities, PermissionItemType } from "../../types/permissions";
import {
  folderPermissionsPath,
  libraryPermissionsPath,
  listItemPermissionsPath,
} from "./effective-permissions-path";
import { toItemCapabilities } from "./parse-base-permissions";

/**
 * Đọc EffectiveBasePermissions (SharePoint REST) — user hiện tại.
 * File/folder theo UniqueId; library theo listId.
 */
export class PermissionService {
  constructor(
    private readonly rest: SharePointRestClient,
    private readonly getLibrary: () => Promise<LibraryContext>,
  ) {}

  /** Gate coarse khi mở thư viện (trước khi có folder UniqueId). */
  async getLibraryCapabilities(options?: { signal?: AbortSignal }): Promise<ItemCapabilities> {
    const library = await this.getLibrary();
    return this.fetchCapabilities(libraryPermissionsPath(library.listId), options?.signal);
  }

  /** Folder theo UniqueId; root → quyền library (không có list item). */
  async getFolderCapabilities(
    uniqueId: string,
    options?: { signal?: AbortSignal },
  ): Promise<ItemCapabilities> {
    const library = await this.getLibrary();
    return this.fetchCapabilities(
      folderPermissionsPath(library.listId, uniqueId, library.rootFolderUniqueId),
      options?.signal,
    );
  }

  async getFileCapabilities(
    uniqueId: string,
    options?: { signal?: AbortSignal },
  ): Promise<ItemCapabilities> {
    const library = await this.getLibrary();
    return this.fetchCapabilities(
      listItemPermissionsPath(library.listId, uniqueId),
      options?.signal,
    );
  }

  async getItemCapabilities(
    type: PermissionItemType,
    uniqueId: string,
    options?: { signal?: AbortSignal },
  ): Promise<ItemCapabilities> {
    return type === "folder"
      ? this.getFolderCapabilities(uniqueId, options)
      : this.getFileCapabilities(uniqueId, options);
  }

  private async fetchCapabilities(path: string, signal?: AbortSignal): Promise<ItemCapabilities> {
    const permissions = await this.rest.get<EffectiveBasePermissionsDto>(path, { signal });
    return toItemCapabilities(permissions);
  }
}
