import { useMutation, useQueryClient } from "@tanstack/react-query";
import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import { useSharePoint } from "../provider/context";
import { getWriteErrorMessage } from "./getWriteErrorMessage";
import { childrenInvalidateKey } from "./childrenInvalidateKey";

/** Tham số upload một file vào folder hiện tại. */
export interface UploadFileInput {
  file: File;
  overwrite?: boolean;
}

/**
 * POST upload file rồi invalidate listChildren folder hiện tại.
 * Toast 1 lần: bắt đầu → cập nhật thành công / lỗi trên cùng id.
 */
export function useUploadFile(parentFolderId: string) {
  const { client, locale, messages, notify } = useSharePoint();
  const queryClient = useQueryClient();

  return useMutation<SharePointItem, Error, UploadFileInput, { toastId: string }>({
    mutationFn: ({ file, overwrite }: UploadFileInput): Promise<SharePointItem> =>
      client.fileUpload.upload(parentFolderId, file, {
        fileName: file.name,
        overwrite,
      }),
    onMutate: ({ file }) => {
      const toastId = notify.info(messages.uploading, file.name);
      return { toastId };
    },
    onSuccess: (_item, { file }, context) => {
      if (!context) return;
      notify.update(context.toastId, {
        intent: "success",
        title: messages.uploadSuccess,
        subtitle: file.name,
      });
    },
    onError: (error, { file }, context) => {
      if (!context) return;
      notify.update(context.toastId, {
        intent: "error",
        title: messages.uploadError,
        subtitle: getWriteErrorMessage(error, messages, messages.uploadError),
      });
    },
    onSettled: async () => {
      await queryClient.invalidateQueries({
        queryKey: childrenInvalidateKey(client, locale, parentFolderId),
      });
    },
  });
}
