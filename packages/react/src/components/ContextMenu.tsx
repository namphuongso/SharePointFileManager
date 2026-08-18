import { useEffect, useRef } from "react";
import {
  type FeatureConfig,
  type SharePointItem,
} from "@namphuongso/sharepoint-file-manager-core";
import {
  MenuDivider,
  MenuItem,
  MenuList,
  tokens,
} from "@fluentui/react-components";
import {
  ArrowDownloadRegular,
  ArrowMoveRegular,
  CopyRegular,
  DeleteRegular,
  DocumentRegular,
  HistoryRegular,
  InfoRegular,
  LockClosedRegular,
  LockOpenRegular,
  OpenRegular,
  PersonSettingsRegular,
  RenameRegular,
  ShareRegular,
} from "@fluentui/react-icons";
import type { Messages } from "../i18n/messages";

export type FileAction =
  | "open"
  | "openInSharePoint"
  | "download"
  | "preview"
  | "rename"
  | "copy"
  | "move"
  | "share"
  | "manageAccess"
  | "properties"
  | "checkout"
  | "checkin"
  | "discardCheckout"
  | "delete"
  | "versionHistory"
  | "activity";

const ACTION_FEATURE: Record<FileAction, keyof FeatureConfig | undefined> = {
  open: undefined,
  openInSharePoint: "openInSharePoint",
  download: "download",
  preview: "preview",
  rename: "rename",
  copy: "copy",
  move: "move",
  share: "share",
  manageAccess: "manageAccess",
  properties: "properties",
  checkout: "checkout",
  checkin: "checkout",
  discardCheckout: "checkout",
  delete: "delete",
  versionHistory: "versionHistory",
  activity: "activityLog",
};

function canShowAction(item: SharePointItem, action: FileAction): boolean {
  const checkedOut = item.capabilities?.isCheckedOut === true;
  switch (action) {
    case "rename":
      return item.capabilities?.canRename !== false;
    case "delete":
      return item.capabilities?.canDelete !== false;
    case "download":
      return item.capabilities?.canDownload !== false;
    case "move":
      return item.capabilities?.canMove !== false;
    case "copy":
      return item.capabilities?.canCopy !== false;
    case "share":
      return item.capabilities?.canShare !== false;
    case "checkout":
      return item.type === "file" && !checkedOut && item.capabilities?.canCheckout !== false;
    case "checkin":
      return item.type === "file" && checkedOut && item.capabilities?.canCheckin !== false;
    case "discardCheckout":
      return item.type === "file" && checkedOut && item.capabilities?.canDiscardCheckout !== false;
    default:
      return true;
  }
}

type MenuAction = { id: FileAction; label: string; show: boolean; dividerBefore?: boolean };

function getActionIcon(action: FileAction) {
  switch (action) {
    case "open":
      return <OpenRegular />;
    case "openInSharePoint":
      return <OpenRegular />;
    case "download":
      return <ArrowDownloadRegular />;
    case "preview":
      return <DocumentRegular />;
    case "rename":
      return <RenameRegular />;
    case "copy":
      return <CopyRegular />;
    case "move":
      return <ArrowMoveRegular />;
    case "share":
      return <ShareRegular />;
    case "manageAccess":
      return <PersonSettingsRegular />;
    case "properties":
      return <InfoRegular />;
    case "checkout":
      return <LockClosedRegular />;
    case "checkin":
      return <LockOpenRegular />;
    case "discardCheckout":
      return <LockOpenRegular />;
    case "delete":
      return <DeleteRegular />;
    case "versionHistory":
      return <HistoryRegular />;
    case "activity":
      return <HistoryRegular />;
    default:
      return undefined;
  }
}

export function ContextMenu({
  item,
  x,
  y,
  messages,
  features,
  onAction,
  onClose,
}: {
  item: SharePointItem;
  x: number;
  y: number;
  messages: Messages;
  features: Required<FeatureConfig>;
  onAction: (action: FileAction) => void;
  onClose: () => void;
}) {
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    function handlePointer(event: MouseEvent) {
      if (ref.current && !ref.current.contains(event.target as Node)) onClose();
    }
    function handleKey(event: KeyboardEvent) {
      if (event.key === "Escape") onClose();
    }
    document.addEventListener("mousedown", handlePointer);
    document.addEventListener("keydown", handleKey);
    return () => {
      document.removeEventListener("mousedown", handlePointer);
      document.removeEventListener("keydown", handleKey);
    };
  }, [onClose]);

  useEffect(() => {
    const firstItem = ref.current?.querySelector<HTMLElement>('[role="menuitem"]');
    firstItem?.focus();
  }, []);

  const actions: MenuAction[] = [
    { id: "open", label: messages.open, show: true },
    { id: "share", label: messages.share, show: true },
    { id: "openInSharePoint", label: messages.openInSharePoint, show: Boolean(item.webUrl) },
    { id: "download", label: messages.download, show: item.type === "file" },
    { id: "delete", label: messages.delete, show: true, dividerBefore: true },
    { id: "preview", label: messages.preview, show: item.type === "file" },
    { id: "rename", label: messages.rename, show: true },
    { id: "move", label: messages.move, show: true },
    { id: "copy", label: messages.copy, show: true },
    { id: "versionHistory", label: messages.versionHistory, show: item.type === "file" },
    { id: "manageAccess", label: messages.manageAccess, show: true },
    { id: "properties", label: messages.properties, show: true, dividerBefore: true },
    { id: "activity", label: messages.activity, show: true },
    { id: "checkout", label: messages.checkout, show: canShowAction(item, "checkout"), dividerBefore: true },
    { id: "checkin", label: messages.checkin, show: canShowAction(item, "checkin") },
    { id: "discardCheckout", label: messages.discardCheckout, show: canShowAction(item, "discardCheckout") },
  ];

  const visible = actions.filter(
    (action) =>
      action.show &&
      canShowAction(item, action.id) &&
      (ACTION_FEATURE[action.id] ? features[ACTION_FEATURE[action.id]!] : true),
  );
  const viewportWidth = typeof window === "undefined" ? x + 248 : window.innerWidth;
  const viewportHeight = typeof window === "undefined" ? y + 320 : window.innerHeight;
  const menuWidth = 240;
  const menuHeight = Math.min(visible.length * 36 + 12, viewportHeight - 24);
  const left = Math.max(8, Math.min(x, viewportWidth - menuWidth - 8));
  const top = Math.max(8, Math.min(y, viewportHeight - menuHeight - 8));

  return (
    <div
      ref={ref}
      role="menu"
      aria-label={item.name}
      tabIndex={-1}
      className="spm-context-menu spm-fixed spm-z-50 spm-rounded-sm spm-border spm-border-sp-border spm-bg-white spm-py-1 spm-shadow-lg"
      style={{ left, top, maxHeight: menuHeight }}
    >
      <MenuList>
        {visible.map((action) => (
          <div key={action.id}>
            {action.dividerBefore ? <MenuDivider /> : null}
            <MenuItem
              icon={getActionIcon(action.id)}
              className={action.id === "delete" ? "spm-context-menu-delete" : undefined}
              style={
                action.id === "delete"
                  ? { color: tokens.colorPaletteRedForeground1 }
                  : undefined
              }
              onClick={() => onAction(action.id)}
            >
              {action.label}
            </MenuItem>
          </div>
        ))}
      </MenuList>
    </div>
  );
}
