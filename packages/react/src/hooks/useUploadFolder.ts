import { useMutation, useQueryClient } from "@tanstack/react-query";
import { useSharePoint } from "../provider/context";
import { getWriteErrorMessage } from "./getWriteErrorMessage";
import { childrenInvalidateKey } from "./childrenInvalidateKey";

/**
 * Upload cây thư mục (input webkitdirectory): ensure folder rồi upload từng file.
 * Toast 1 lần: bắt đầu → cập nhật thành công / lỗi trên cùng id.
 */
export function useUploadFolder(parentFolderId: string) {
  const { client, locale, messages, notify } = useSharePoint();
  const queryClient = useQueryClient();

  return useMutation<void, Error, File[], { toastId: string; topFolder: string }>({
    mutationFn: async (files: File[]) => {
      const folderPaths = new Set<string>();
      for (const file of files) {
        const relative = file.webkitRelativePath || file.name;
        const parts = relative.split("/").filter(Boolean);
        parts.pop();
        let acc = "";
        for (const part of parts) {
          acc = acc ? `${acc}/${part}` : part;
          folderPaths.add(acc);
        }
      }

      const sorted = [...folderPaths].sort(
        (a, b) => a.split("/").length - b.split("/").length || a.localeCompare(b),
      );
      for (const path of sorted) {
        await client.folderCreate.ensureRelativePath(parentFolderId, path);
      }

      for (const file of files) {
        await client.fileUpload.uploadRelativeFile(parentFolderId, file);
      }
    },
    onMutate: (files) => {
      const first = files[0];
      const topFolder = first
        ? (first.webkitRelativePath || first.name).split("/")[0] ?? ""
        : "";
      const toastId = notify.info(messages.uploading, topFolder || undefined);
      return { toastId, topFolder };
    },
    onSuccess: (_data, _files, context) => {
      if (!context) return;
      notify.update(context.toastId, {
        intent: "success",
        title: messages.uploadFolderSuccess,
        subtitle: context.topFolder || undefined,
      });
    },
    onError: (error, _files, context) => {
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
