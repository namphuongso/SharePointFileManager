import { useState, type ReactElement } from "react";
import {
  Menu,
  MenuPopover,
  MenuTrigger,
  Button,
} from "@fluentui/react-components";
import { AddRegular } from "@fluentui/react-icons";
import type { Messages } from "../../types/messages";
import type { NewItemAction } from "./newItemActions";
import { NewItemMenuList } from "./NewItemMenuList";
import { useFileManagerStyles } from "./useFileManagerStyles";

export interface NewItemToolbarMenuProps {
  messages: Messages;
  disabled?: boolean;
  onAction: (action: NewItemAction) => void;
  onPickFiles: () => void;
  onPickFolder: () => void;
}

/** Nút + menu New trên toolbar (pill icon). Mở thẳng danh sách tạo mới. */
export function NewItemToolbarMenu({
  messages,
  disabled,
  onAction,
  onPickFiles,
  onPickFolder,
}: NewItemToolbarMenuProps): ReactElement {
  const styles = useFileManagerStyles();
  const [open, setOpen] = useState(false);

  return (
    <Menu open={open} onOpenChange={(_, data) => setOpen(data.open)}>
      <MenuTrigger disableButtonEnhancement>
        <Button
          appearance="subtle"
          shape="circular"
          className={styles.commandIconButton}
          icon={<AddRegular fontSize={20} />}
          disabled={disabled}
          aria-label={messages.newItem}
          title={messages.newItem}
        />
      </MenuTrigger>
      <MenuPopover>
        <NewItemMenuList
          messages={messages}
          disabled={disabled}
          onAction={(action) => {
            onAction(action);
            setOpen(false);
          }}
          onPickFiles={() => {
            // persistOnClick: picker mở trước khi đóng menu (giữ user gesture).
            onPickFiles();
            setOpen(false);
          }}
          onPickFolder={() => {
            onPickFolder();
            setOpen(false);
          }}
        />
      </MenuPopover>
    </Menu>
  );
}
