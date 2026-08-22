import type { SharePointItemType } from "../types/models";

/**
 * UniqueId từ REST bắt buộc: SharePointItem.id dùng làm key UI / breadcrumb.
 * REST để UniqueId optional; thiếu thì fail sớm thay vì id rỗng.
 */
export function requireUniqueId(id: string | undefined, kind: SharePointItemType): string {
  if (!id) throw new Error(`SharePoint ${kind} is missing UniqueId`);
  return id;
}
