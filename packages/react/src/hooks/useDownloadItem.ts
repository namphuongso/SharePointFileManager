import { useCallback, useState } from "react";
import {
  isSharePointError,
  SharePointErrorCode,
  type FolderZipProgress,
  type SharePointItem,
} from "@namphuongso/sharepoint-file-manager-core";
import { useSharePoint } from "../provider/context";
import type { Messages } from "../types/messages";
import { getErrorMessage } from "./getErrorMessage";
import { triggerBrowserDownload } from "./triggerBrowserDownload";

/** Kết quả tải: đã lưu, bị từ chối OpenItems, hoặc lỗi. */
export type DownloadItemResult =
  | { status: "downloaded" }
  | { status: "denied" }
  | { status: "error"; message: string };

function folderProgressSubtitle(progress: FolderZipProgress, messages: Messages): string {
  if (progress.phase === "listing") return messages.downloadingFolderListing;
  if (progress.phase === "zipping") return messages.downloadingFolderZipping;
  return messages.downloadingFolderFiles
    .replace("{done}", String(progress.done))
    .replace("{total}", String(progress.total));
}

/**
 * Tải file (/$value) hoặc folder (.zip đệ quy) khi user chọn Tải xuống.
 * Toast `progress` sticky đến khi xong — folder cập nhật phase listing → files → zip.
 */
export function useDownloadItem() {
  const { client, messages, notify } = useSharePoint();
  const [downloadingId, setDownloadingId] = useState<string>();

  const downloadItem = useCallback(
    async (item: SharePointItem): Promise<DownloadItemResult> => {
      if (item.type !== "file" && item.type !== "folder") {
        return { status: "error", message: "Not a file or folder" };
      }
      setDownloadingId(item.id);
      const toastId = notify.progress(
        messages.downloading,
        item.type === "folder" ? `${item.name}.zip` : item.name,
      );
      try {
        const caps = await client.permissions.getItemCapabilities(item.type, item.id);
        if (!caps.canOpen) {
          notify.dismiss(toastId);
          return { status: "denied" };
        }

        const result =
          item.type === "folder"
            ? await client.files.downloadFolder(item.id, {
                folderName: item.name,
                onProgress: (progress) => {
                  notify.update(toastId, {
                    intent: "info",
                    title: messages.downloading,
                    subtitle: `${folderProgressSubtitle(progress, messages)} — ${item.name}`,
                  });
                },
              })
            : await client.files.download(item.id);
        const saveName =
          result.fileName ?? (item.type === "folder" ? `${item.name}.zip` : item.name);
        triggerBrowserDownload(result.blob, saveName);
        notify.update(toastId, {
          intent: "success",
          title: messages.downloadSuccess,
          subtitle: saveName,
        });
        return { status: "downloaded" };
      } catch (error) {
        if (
          isSharePointError(error) &&
          error.code === SharePointErrorCode.Forbidden
        ) {
          notify.dismiss(toastId);
          return { status: "denied" };
        }
        const message = getErrorMessage(error);
        notify.update(toastId, {
          intent: "error",
          title: messages.downloadError,
          subtitle: message,
        });
        return { status: "error", message };
      } finally {
        setDownloadingId(undefined);
      }
    },
    [client, messages, notify],
  );

  return { downloadItem, downloadingId };
}
