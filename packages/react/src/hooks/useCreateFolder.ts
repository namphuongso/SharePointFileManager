import { useMutation, useQueryClient } from "@tanstack/react-query";
import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import { useSharePoint } from "../provider/context";
import { getWriteErrorMessage } from "./getWriteErrorMessage";
import { childrenInvalidateKey } from "./childrenInvalidateKey";

/**
 * POST tạo thư mục con rồi invalidate listChildren folder hiện tại.
 * Toast thành công / lỗi kèm tên folder mới.
 */
export function useCreateFolder(parentFolderId: string) {
  const { client, locale, messages, notify } = useSharePoint();
  const queryClient = useQueryClient();

  return useMutation<SharePointItem, Error, string, { toastId: string }>({
    mutationFn: (name: string): Promise<SharePointItem> =>
      client.folderCreate.create(parentFolderId, name),
    onMutate: (name) => {
      const toastId = notify.info(messages.uploading, name);
      return { toastId };
    },
    onSuccess: (_item, name, context) => {
      if (!context) return;
      notify.update(context.toastId, {
        intent: "success",
        title: messages.createFolderSuccess,
        subtitle: name,
      });
    },
    onError: (error, name, context) => {
      if (!context) return;
      notify.update(context.toastId, {
        intent: "error",
        title: messages.createFolderError,
        subtitle: getWriteErrorMessage(error, messages, messages.createFolderError),
      });
    },
    onSettled: async () => {
      await queryClient.invalidateQueries({
        queryKey: childrenInvalidateKey(client, locale, parentFolderId),
      });
    },
  });
}
