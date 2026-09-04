import { ArrowUploadRegular, FolderArrowUpRegular } from "@fluentui/react-icons";
import {
  FileTypeIcon as FluentFileTypeIcon,
  FileIconType,
} from "@fluentui/react-icons-file-type";
import type { NewDocumentKind } from "@namphuongso/sharepoint-file-manager-core";
import {
  MenuDivider,
  MenuItem,
  MenuList,
} from "@fluentui/react-components";
import type { Messages } from "../../types/messages";
import type { NewItemAction } from "./newItemActions";

export interface NewItemMenuListProps {
  messages: Messages;
  disabled?: boolean;
  onAction: (action: NewItemAction) => void;
  /** Mở OS file picker — gọi đồng bộ trong click (input nằm ngoài Menu). */
  onPickFiles: () => void;
  onPickFolder: () => void;
}

/** Đuôi file của từng loại tài liệu — khớp extension SharePoint sẽ tạo. */
function extensionForKind(kind: NewDocumentKind): string {
  const extensionByKind = {
    word: "docx",
    excel: "xlsx",
    powerpoint: "pptx",
  } as const satisfies Record<NewDocumentKind, string>;
  return extensionByKind[kind];
}

function newItemIcon(action: NewItemAction): JSX.Element {
  if (action.type === "uploadFiles") return <ArrowUploadRegular />;
  if (action.type === "uploadFolder") return <FolderArrowUpRegular />;
  if (action.type === "folder") {
    return <FluentFileTypeIcon type={FileIconType.folder} size={20} className="spm-shrink-0" />;
  }
  return (
    <FluentFileTypeIcon
      extension={extensionForKind(action.kind)}
      size={20}
      className="spm-shrink-0"
    />
  );
}

/**
 * Nội dung menu New. Upload dùng persistOnClick + showPicker đồng bộ — Fluent
 * mặc định đóng menu trước onClick (mất user gesture → picker không mở).
 */
export function NewItemMenuList({
  messages,
  disabled,
  onAction,
  onPickFiles,
  onPickFolder,
}: NewItemMenuListProps) {
  return (
    <MenuList>
      <MenuItem
        icon={newItemIcon({ type: "folder" })}
        disabled={disabled}
        onClick={() => onAction({ type: "folder" })}
      >
        {messages.newFolder}
      </MenuItem>
      <MenuItem
        icon={newItemIcon({ type: "uploadFiles" })}
        disabled={disabled}
        persistOnClick
        onClick={() => onPickFiles()}
      >
        {messages.uploadFiles}
      </MenuItem>
      <MenuItem
        icon={newItemIcon({ type: "uploadFolder" })}
        disabled={disabled}
        persistOnClick
        onClick={() => onPickFolder()}
      >
        {messages.uploadFolder}
      </MenuItem>
      <MenuDivider />
      <MenuItem
        icon={newItemIcon({ type: "document", kind: "word" })}
        disabled={disabled}
        onClick={() => onAction({ type: "document", kind: "word" })}
      >
        {messages.newWord}
      </MenuItem>
      <MenuItem
        icon={newItemIcon({ type: "document", kind: "excel" })}
        disabled={disabled}
        onClick={() => onAction({ type: "document", kind: "excel" })}
      >
        {messages.newExcel}
      </MenuItem>
      <MenuItem
        icon={newItemIcon({ type: "document", kind: "powerpoint" })}
        disabled={disabled}
        onClick={() => onAction({ type: "document", kind: "powerpoint" })}
      >
        {messages.newPowerPoint}
      </MenuItem>
    </MenuList>
  );
}
