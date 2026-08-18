import type { FeatureConfig, SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../i18n/messages";
import { Button } from "./ui";

export type SelectionAction =
  | "download"
  | "copy"
  | "move"
  | "share"
  | "rename"
  | "preview"
  | "manageAccess"
  | "bulkMetadata"
  | "delete";

export function SelectionToolbar({
  items,
  messages,
  features,
  onClear,
  onAction,
}: {
  items: SharePointItem[];
  messages: Messages;
  features: Required<FeatureConfig>;
  onClear: () => void;
  onAction: (action: SelectionAction) => void;
}) {
  if (items.length === 0) return null;

  const single = items[0];
  const fileItems = items.filter((item) => item.type === "file");
  const canDownload = fileItems.length > 0 && features.download;
  const canShare = items.length === 1 && features.share;
  const canRename = items.length === 1 && features.rename;
  const canPreview = items.length === 1 && single?.type === "file" && features.preview;
  const canManageAccess = items.length === 1 && features.manageAccess;

  return (
    <div className="spm-flex spm-flex-wrap spm-items-center spm-gap-2 spm-border-b spm-border-sp-border spm-bg-slate-50 spm-px-3 spm-py-2">
      <span className="spm-text-sm">
        {items.length} {messages.selected}
      </span>
      {canDownload ? (
        <Button onClick={() => onAction("download")}>
          {messages.download}
          {fileItems.length > 1 ? ` (${fileItems.length})` : ""}
        </Button>
      ) : null}
      {canPreview ? <Button onClick={() => onAction("preview")}>{messages.preview}</Button> : null}
      {canRename ? <Button onClick={() => onAction("rename")}>{messages.rename}</Button> : null}
      {features.copy ? <Button onClick={() => onAction("copy")}>{messages.copy}</Button> : null}
      {features.move ? <Button onClick={() => onAction("move")}>{messages.move}</Button> : null}
      {canShare ? <Button onClick={() => onAction("share")}>{messages.share}</Button> : null}
      {canManageAccess ? <Button onClick={() => onAction("manageAccess")}>{messages.manageAccess}</Button> : null}
      {features.bulkMetadata && features.metadata && items.length > 1 ? (
        <Button onClick={() => onAction("bulkMetadata")}>{messages.bulkEditMetadata}</Button>
      ) : null}
      {features.delete ? (
        <Button variant="danger" onClick={() => onAction("delete")}>
          {messages.delete}
        </Button>
      ) : null}
      <Button onClick={onClear}>{messages.cancel}</Button>
    </div>
  );
}
