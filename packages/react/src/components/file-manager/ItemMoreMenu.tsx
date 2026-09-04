import { useState, type MouseEvent, type ReactElement } from "react";
import {
  Button,
  Menu,
  MenuPopover,
  MenuTrigger,
  mergeClasses,
} from "@fluentui/react-components";
import { MoreHorizontalRegular } from "@fluentui/react-icons";
import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../../types/messages";
import { ItemActionsMenuList } from "./ItemActionsMenuList";
import { useFileManagerStyles } from "./useFileManagerStyles";

export interface ItemMoreMenuProps {
  item: SharePointItem;
  messages: Messages;
  /** Hiện nút khi hover dòng hoặc menu đang mở. */
  visible: boolean;
  busy?: boolean;
  onOpen: (item: SharePointItem) => void;
  onDownload: (item: SharePointItem) => void;
  onDelete: (item: SharePointItem) => void;
}

/**
 * Nút ⋯ cuối cột đầu tiên — menu Mở / Tải xuống / Xóa.
 * stopPropagation để không kích hoạt click mở dòng.
 */
export function ItemMoreMenu({
  item,
  messages,
  visible,
  busy,
  onOpen,
  onDownload,
  onDelete,
}: ItemMoreMenuProps): ReactElement {
  const styles = useFileManagerStyles();
  const [open, setOpen] = useState(false);
  const show = visible || open;

  function stopRowClick(event: MouseEvent) {
    event.preventDefault();
    event.stopPropagation();
  }

  return (
    <Menu
      open={open}
      onOpenChange={(_, data) => setOpen(data.open)}
      positioning={{ position: "below", align: "end" }}
    >
      <MenuTrigger disableButtonEnhancement>
        <Button
          appearance="subtle"
          shape="circular"
          className={mergeClasses(
            styles.rowMoreButton,
            show && styles.rowMoreButtonVisible,
          )}
          icon={<MoreHorizontalRegular fontSize={20} />}
          aria-label={messages.moreActions}
          title={messages.moreActions}
          disabled={busy}
          onClick={stopRowClick}
          onMouseDown={stopRowClick}
        />
      </MenuTrigger>
      <MenuPopover onClick={stopRowClick}>
        <ItemActionsMenuList
          item={item}
          messages={messages}
          menuOpen={open}
          busy={busy}
          onOpen={() => {
            setOpen(false);
            onOpen(item);
          }}
          onDownload={() => {
            setOpen(false);
            onDownload(item);
          }}
          onDelete={() => {
            setOpen(false);
            onDelete(item);
          }}
        />
      </MenuPopover>
    </Menu>
  );
}
