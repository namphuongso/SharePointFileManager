import type { ReactElement } from "react";
import { MenuItem, MenuList } from "@fluentui/react-components";
import { ArrowDownloadRegular, OpenRegular } from "@fluentui/react-icons";
import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import { useItemOpenCapability } from "../../hooks/useItemOpenCapability";
import type { Messages } from "../../types/messages";

export interface ItemActionsMenuListProps {
  item: SharePointItem;
  messages: Messages;
  /** Menu đang mở — mới GET EffectiveBasePermissions (lazy). */
  menuOpen: boolean;
  busy?: boolean;
  onOpen: () => void;
  onDownload: () => void;
}

/**
 * Mục menu thao tác: Mở + Tải xuống.
 * Không đủ quyền → ẩn (không disabled). Đang GET quyền → danh sách trống tạm.
 */
export function ItemActionsMenuList({
  item,
  messages,
  menuOpen,
  busy,
  onOpen,
  onDownload,
}: ItemActionsMenuListProps): ReactElement {
  const isFile = item.type === "file";
  const { canOpen, canView, isLoading } = useItemOpenCapability(item, menuOpen);

  const showOpen = !isLoading && (isFile ? canOpen : canView);
  const showDownload = !isLoading && canOpen;

  return (
    <MenuList>
      {showOpen ? (
        <MenuItem icon={<OpenRegular />} disabled={busy} onClick={onOpen}>
          {isFile ? messages.openFile : messages.openFolder}
        </MenuItem>
      ) : null}
      {showDownload ? (
        <MenuItem icon={<ArrowDownloadRegular />} disabled={busy} onClick={onDownload}>
          {messages.download}
        </MenuItem>
      ) : null}
    </MenuList>
  );
}
