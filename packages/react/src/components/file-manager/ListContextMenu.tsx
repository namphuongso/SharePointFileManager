import type { ReactElement } from "react";
import {
  Menu,
  MenuDivider,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
} from "@fluentui/react-components";
import { AddRegular, ArrowClockwiseRegular } from "@fluentui/react-icons";
import type { Messages } from "../../types/messages";
import type { NewItemAction } from "./newItemActions";
import { NewItemMenuList } from "./NewItemMenuList";

export interface ListContextMenuProps {
  open: boolean;
  /** Điểm neo ảo tại vị trí chuột phải. */
  anchor: { x: number; y: number } | null;
  onOpenChange: (open: boolean) => void;
  messages: Messages;
  /** Hiện submenu Mới khi có AddListItems. */
  canAdd?: boolean;
  /** Đang ghi — khoá mục Mới tạm thời. */
  writeBusy?: boolean;
  onAction: (action: NewItemAction) => void;
  onPickFiles: () => void;
  onPickFolder: () => void;
  /** Làm mới danh sách (cùng nút refresh toolbar). */
  onRefresh: () => void;
}

/**
 * Menu chuột phải trên vùng danh sách — Mới (nếu có quyền) + Làm mới.
 */
export function ListContextMenu({
  open,
  anchor,
  onOpenChange,
  messages,
  canAdd,
  writeBusy,
  onAction,
  onPickFiles,
  onPickFolder,
  onRefresh,
}: ListContextMenuProps): ReactElement | null {
  if (!anchor) return null;

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
        <MenuList>
          {canAdd ? (
            <>
              <Menu>
                <MenuTrigger disableButtonEnhancement>
                  <MenuItem icon={<AddRegular />} disabled={writeBusy}>
                    {messages.newItem}
                  </MenuItem>
                </MenuTrigger>
                <MenuPopover>
                  <NewItemMenuList
                    messages={messages}
                    disabled={writeBusy}
                    onPickFiles={() => closeAfter(onPickFiles)}
                    onPickFolder={() => closeAfter(onPickFolder)}
                    onAction={(action) => closeAfter(() => onAction(action))}
                  />
                </MenuPopover>
              </Menu>
              <MenuDivider />
            </>
          ) : null}
          <MenuItem
            icon={<ArrowClockwiseRegular />}
            onClick={() => closeAfter(onRefresh)}
          >
            {messages.refresh}
          </MenuItem>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}
