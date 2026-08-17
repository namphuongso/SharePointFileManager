import type { FeatureConfig, SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../i18n/messages";

export type FileAction =
  | "open"
  | "download"
  | "preview"
  | "rename"
  | "copy"
  | "move"
  | "share"
  | "manageAccess"
  | "delete"
  | "versionHistory";

const ACTION_FEATURE: Record<FileAction, keyof FeatureConfig | undefined> = {
  open: "openInSharePoint",
  download: "download",
  preview: "preview",
  rename: "rename",
  copy: "copy",
  move: "move",
  share: "share",
  manageAccess: "manageAccess",
  delete: "delete",
  versionHistory: "versionHistory",
};

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
  const actions: Array<{ id: FileAction; label: string; show: boolean }> = [
    { id: "open", label: messages.open, show: true },
    { id: "download", label: messages.download, show: item.type === "file" },
    { id: "preview", label: messages.preview, show: item.type === "file" },
    { id: "rename", label: messages.rename, show: true },
    { id: "copy", label: messages.copy, show: true },
    { id: "move", label: messages.move, show: true },
    { id: "share", label: messages.share, show: true },
    { id: "manageAccess", label: messages.manageAccess, show: true },
    { id: "versionHistory", label: messages.versionHistory, show: item.type === "file" },
    { id: "delete", label: messages.delete, show: true },
  ];

  return (
    <div
      className="spm-fixed spm-z-50 spm-min-w-[180px] spm-rounded-md spm-border spm-border-sp-border spm-bg-white spm-py-1 spm-shadow-lg"
      style={{ left: x, top: y }}
      onMouseLeave={onClose}
    >
      {actions
        .filter((action) => action.show && (ACTION_FEATURE[action.id] ? features[ACTION_FEATURE[action.id]!] : true))
        .map((action) => (
          <button
            type="button"
            key={action.id}
            className={`spm-block spm-w-full spm-px-3 spm-py-1.5 spm-text-left spm-text-sm hover:spm-bg-slate-50 ${
              action.id === "delete" ? "spm-text-sp-danger" : ""
            }`}
            onClick={() => onAction(action.id)}
          >
            {action.label}
          </button>
        ))}
    </div>
  );
}
