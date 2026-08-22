import type { SharePointItem, SharePointItemType } from "../types/models";

/**
 * Field chung SP.File / SP.Folder → SharePointItem (cùng shape cho UI).
 * `uniqueId` phải đã qua requireUniqueId: REST để UniqueId optional, model thì id là key.
 */
export function mapRestItem(
  type: SharePointItemType,
  name: string,
  uniqueId: string,
  lastModifiedDateTime?: string,
  size?: number,
): SharePointItem {
  return { id: uniqueId, name, type, size, lastModifiedDateTime };
}
