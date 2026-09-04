import { useMutation, useQueryClient } from "@tanstack/react-query";
import type {
  CreateDocumentOptions,
  NewDocumentKind,
  SharePointItem,
} from "@namphuongso/sharepoint-file-manager-core";
import { useSharePoint } from "../provider/context";
import { getWriteErrorMessage } from "./getWriteErrorMessage";
import { childrenInvalidateKey } from "./childrenInvalidateKey";

export interface CreateDocumentInput {
  kind: NewDocumentKind;
  name: string;
  overwrite?: boolean;
}

/**
 * POST tạo file trống (Word / Excel / PowerPoint) rồi invalidate listChildren.
 * Toast thành công / lỗi kèm tên file mới.
 */
export function useCreateDocument(parentFolderId: string) {
  const { client, locale, messages, notify } = useSharePoint();
  const queryClient = useQueryClient();

  return useMutation<SharePointItem, Error, CreateDocumentInput, { toastId: string }>({
    mutationFn: ({ kind, name, overwrite }: CreateDocumentInput): Promise<SharePointItem> => {
      const options: CreateDocumentOptions = { name, overwrite };
      return client.documentCreate.create(parentFolderId, kind, options);
    },
    onMutate: ({ name }) => {
      const toastId = notify.info(messages.uploading, name);
      return { toastId };
    },
    onSuccess: (_item, { name }, context) => {
      if (!context) return;
      notify.update(context.toastId, {
        intent: "success",
        title: messages.createDocumentSuccess,
        subtitle: name,
      });
    },
    onError: (error, { name }, context) => {
      if (!context) return;
      notify.update(context.toastId, {
        intent: "error",
        title: messages.createDocumentError,
        subtitle: getWriteErrorMessage(error, messages, messages.createDocumentError),
      });
    },
    onSettled: async () => {
      await queryClient.invalidateQueries({
        queryKey: childrenInvalidateKey(client, locale, parentFolderId),
      });
    },
  });
}
