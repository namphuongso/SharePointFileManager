import { useQuery } from "@tanstack/react-query";
import {
  isSharePointError,
  SharePointErrorCode,
  type SharePointItem,
} from "@namphuongso/sharepoint-file-manager-core";
import { useSharePoint } from "../provider/context";
import { queryKeys } from "./queryKeys";

/** Quyền mở / tải / xóa trên một item — lazy khi mở menu (không GET theo từng dòng list). */
export interface ItemOpenAccess {
  isLoading: boolean;
  /** OpenItems — mở file / tải file·folder. */
  canOpen: boolean;
  /** ViewListItems — mở (vào) folder. */
  canView: boolean;
  /** DeleteListItems — soft-delete (recycle). */
  canDelete: boolean;
}

/**
 * GET EffectiveBasePermissions theo UniqueId khi `enabled` (menu đang mở).
 * Cache 60s — cùng item mở lại menu không GET lại ngay.
 */
export function useItemOpenCapability(
  item: SharePointItem | null | undefined,
  enabled: boolean,
): ItemOpenAccess {
  const { client } = useSharePoint();
  const query = useQuery({
    queryKey: queryKeys.itemCapabilities(
      client.config.siteId,
      client.cacheScope,
      item?.type ?? "",
      item?.id ?? "",
    ),
    enabled: Boolean(item) && enabled,
    queryFn: ({ signal }) =>
      client.permissions.getItemCapabilities(item!.type, item!.id, { signal }),
    staleTime: 60_000,
  });

  const forbidden =
    query.isError &&
    isSharePointError(query.error) &&
    query.error.code === SharePointErrorCode.Forbidden;

  const canOpen = !forbidden && query.isSuccess && query.data.canOpen;
  const canView = !forbidden && query.isSuccess && query.data.canView;
  const canDelete = !forbidden && query.isSuccess && query.data.canDelete;

  return {
    isLoading: Boolean(item) && enabled && query.isPending,
    canOpen,
    canView,
    canDelete,
  };
}
