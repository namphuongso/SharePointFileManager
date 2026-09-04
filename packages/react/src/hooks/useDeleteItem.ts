import { useMutation, useQueryClient } from "@tanstack/react-query";
import {
  isSharePointError,
  SharePointError,
  SharePointErrorCode,
  type SharePointItem,
} from "@namphuongso/sharepoint-file-manager-core";
import { useSharePoint } from "../provider/context";
import { childrenInvalidateKey } from "./childrenInvalidateKey";
import { getErrorMessage } from "./getErrorMessage";

/**
 * Soft-delete (recycle) file/folder rồi invalidate listChildren + tab accessible.
 * Toast progress → success/error; Forbidden map sang noDeletePermission.
 */
export function useDeleteItem(parentFolderId: string) {
  const { client, locale, messages, notify } = useSharePoint();
  const queryClient = useQueryClient();

  return useMutation<void, Error, SharePointItem, { toastId: string }>({
    mutationFn: async (item) => {
      const caps = await client.permissions.getItemCapabilities(item.type, item.id);
      if (!caps.canDelete) {
        throw new SharePointError({
          code: SharePointErrorCode.Forbidden,
          message: "Missing DeleteListItems",
        });
      }
      await client.itemDelete.delete(item.type, item.id);
    },
    onMutate: (item) => {
      const toastId = notify.progress(messages.deleting, item.name);
      return { toastId };
    },
    onSuccess: (_void, item, context) => {
      if (!context) return;
      notify.update(context.toastId, {
        intent: "success",
        title: messages.deleteSuccess,
        subtitle: item.name,
      });
    },
    onError: (error, _item, context) => {
      if (!context) return;
      const subtitle =
        isSharePointError(error) && error.code === SharePointErrorCode.Forbidden
          ? messages.noDeletePermission
          : getErrorMessage(error, messages.deleteError);
      notify.update(context.toastId, {
        intent: "error",
        title: messages.deleteError,
        subtitle,
      });
    },
    onSettled: async () => {
      await Promise.all([
        queryClient.invalidateQueries({
          queryKey: childrenInvalidateKey(client, locale, parentFolderId),
        }),
        queryClient.invalidateQueries({
          queryKey: ["sp", client.config.siteId, client.cacheScope, "accessible"],
        }),
      ]);
    },
  });
}

/** True khi mutation lỗi vì thiếu DeleteListItems (banner UI). */
export function isDeleteDenied(error: unknown): boolean {
  return isSharePointError(error) && error.code === SharePointErrorCode.Forbidden;
}
