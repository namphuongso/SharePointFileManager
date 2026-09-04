import type { SharePointClient } from "@namphuongso/sharepoint-file-manager-core";

/**
 * Prefix invalidate listChildren — phải khớp useFolderChildren
 * (`client.config.locale ?? locale`), không chỉ `locale` context.
 */
export function childrenInvalidateKey(
  client: SharePointClient,
  locale: string,
  folderId: string,
) {
  return [
    "sp",
    client.config.siteId,
    client.cacheScope,
    client.config.locale ?? locale,
    "children",
    folderId,
  ] as const;
}
