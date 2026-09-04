import type { ReactElement } from "react";
import { MenuDivider, MenuItem, MenuList } from "@fluentui/react-components";
import { ArrowDownloadRegular, DeleteRegular, OpenRegular } from "@fluentui/react-icons";
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
  onDelete: () => void;
}

/**
 * Mục menu thao tác: Mở + Tải xuống + Xóa.
 * Không đủ quyền → ẩn (không disabled). Đang GET quyền → danh sách trống tạm.
 */
export function ItemActionsMenuList({
  item,
  messages,
  menuOpen,
  busy,
  onOpen,
  onDownload,
  onDelete,
}: ItemActionsMenuListProps): ReactElement {
  const isFile = item.type === "file";
  const { canOpen, canView, canDelete, isLoading } = useItemOpenCapability(item, menuOpen);

  const showOpen = !isLoading && (isFile ? canOpen : canView);
  const showDownload = !isLoading && canOpen;
  const showDelete = !isLoading && canDelete;
  const showDivider = (showOpen || showDownload) && showDelete;

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
      {showDivider ? <MenuDivider /> : null}
      {showDelete ? (
        <MenuItem icon={<DeleteRegular />} disabled={busy} onClick={onDelete}>
          {messages.delete}
        </MenuItem>
      ) : null}
    </MenuList>
  );
}
