import type { ReactElement } from "react";
import {
  Menu,
  MenuPopover,
} from "@fluentui/react-components";
import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../../types/messages";
import { ItemActionsMenuList } from "./ItemActionsMenuList";

export interface ItemContextMenuProps {
  open: boolean;
  /** Điểm neo ảo tại vị trí chuột phải trên dòng. */
  anchor: { x: number; y: number } | null;
  item: SharePointItem | null;
  onOpenChange: (open: boolean) => void;
  messages: Messages;
  /** Đang mở / tải — khoá menu. */
  busy?: boolean;
  onOpen: (item: SharePointItem) => void;
  onDownload: (item: SharePointItem) => void;
}

/**
 * Menu chuột phải trên dòng: Mở + Tải xuống (folder → zip).
 * Quyền lazy khi menu mở (useItemOpenCapability).
 */
export function ItemContextMenu({
  open,
  anchor,
  item,
  onOpenChange,
  messages,
  busy,
  onOpen,
  onDownload,
}: ItemContextMenuProps): ReactElement | null {
  if (!anchor || !item) return null;

  const target = {
    getBoundingClientRect: () =>
      DOMRect.fromRect({
        x: anchor.x,
        y: anchor.y,
        width: 0,
        height: 0,
      }),
  };

  function closeAfter(run: () => void) {
    run();
    onOpenChange(false);
  }

  return (
    <Menu
      open={open}
      onOpenChange={(_, data) => onOpenChange(data.open)}
      positioning={{ target, position: "below", align: "start" }}
    >
      <MenuPopover>
        <ItemActionsMenuList
          item={item}
          messages={messages}
          menuOpen={open}
          busy={busy}
          onOpen={() => closeAfter(() => onOpen(item))}
          onDownload={() => closeAfter(() => onDownload(item))}
        />
      </MenuPopover>
    </Menu>
  );
}
