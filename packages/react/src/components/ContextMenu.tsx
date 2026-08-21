import { useEffect, useRef } from "react";
import {
  type FeatureConfig,
  type SharePointItem,
  canPerformItemAction,
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

function canEnableAction(item: SharePointItem, action: FileAction): boolean {
  return canPerformItemAction(item, action);
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
    { id: "open", label: messages.open, show: item.type === "folder" || canPerformItemAction(item, "open") },
    { id: "share", label: messages.share, show: features.share && canPerformItemAction(item, "share") },
    {
      id: "openInSharePoint",
      label: messages.openInSharePoint,
      show: features.openInSharePoint && canPerformItemAction(item, "openInSharePoint"),
    },
    {
      id: "download",
      label: messages.download,
      show: features.download && item.type === "file" && canPerformItemAction(item, "download"),
    },
    {
      id: "delete",
      label: messages.delete,
      show: features.delete && canPerformItemAction(item, "delete"),
      dividerBefore: true,
    },
    {
      id: "preview",
      label: messages.preview,
      show: features.preview && item.type === "file" && canPerformItemAction(item, "preview"),
    },
    { id: "rename", label: messages.rename, show: features.rename && canPerformItemAction(item, "rename") },
    { id: "move", label: messages.move, show: features.move && canPerformItemAction(item, "move") },
    { id: "copy", label: messages.copy, show: features.copy && canPerformItemAction(item, "copy") },
    {
      id: "versionHistory",
      label: messages.versionHistory,
      show: features.versionHistory && item.type === "file" && canPerformItemAction(item, "versionHistory"),
    },
    {
      id: "manageAccess",
      label: messages.manageAccess,
      show: features.manageAccess && canPerformItemAction(item, "manageAccess"),
    },
    { id: "properties", label: messages.properties, show: features.properties, dividerBefore: true },
    { id: "activity", label: messages.activity, show: features.activityLog && canPerformItemAction(item, "activity") },
    {
      id: "checkout",
      label: messages.checkout,
      show: features.checkout && item.type === "file" && canPerformItemAction(item, "checkout"),
      dividerBefore: true,
    },
    {
      id: "checkin",
      label: messages.checkin,
      show: features.checkout && item.type === "file" && canPerformItemAction(item, "checkin"),
    },
    {
      id: "discardCheckout",
      label: messages.discardCheckout,
      show: features.checkout && item.type === "file" && canPerformItemAction(item, "discardCheckout"),
    },
  ];

  const visible = actions.filter(
    (action) =>
      action.show &&
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
              disabled={!canEnableAction(item, action.id)}
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
