import { useCallback, useState } from "react";
import {
  isSharePointError,
  SharePointErrorCode,
} from "@namphuongso/sharepoint-file-manager-core";
import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import { useSharePoint } from "../provider/context";
import { getErrorMessage } from "./getErrorMessage";

/** Kết quả mở file: URL để mở tab mới, hoặc lý do không mở được. */
export type OpenItemResult =
  | { status: "opened"; url: string }
  | { status: "denied" }
  | { status: "error"; message: string };

/**
 * Mở file theo UniqueId khi user bấm dòng (chỉ đọc, GET một lần khi click).
 * 1. GET EffectiveBasePermissions — không có OpenItems → denied, không GET file.
 * 2. GET ServerRelativeUrl → URL ?web=1 (preview mọi loại file; không preview được thì tải về).
 * Folder không qua hook này — UI tự điều hướng breadcrumb.
 */
export function useOpenItem() {
  const { client } = useSharePoint();
  const [openingId, setOpeningId] = useState<string>();

  const openItem = useCallback(
    async (item: SharePointItem): Promise<OpenItemResult> => {
      if (item.type !== "file") return { status: "error", message: "Not a file" };
      setOpeningId(item.id);
      try {
        const caps = await client.permissions.getItemCapabilities(item.type, item.id);
        if (!caps.canOpen) return { status: "denied" };

        const url = await client.files.getOpenUrl(item.id);
        return { status: "opened", url };
      } catch (error) {
        if (
          isSharePointError(error) &&
          error.code === SharePointErrorCode.Forbidden
        ) {
          return { status: "denied" };
        }
        return { status: "error", message: getErrorMessage(error) };
      } finally {
        setOpeningId(undefined);
      }
    },
    [client],
  );

  return { openItem, openingId };
}
