import { useEffect, useState } from "react";
import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../../i18n/messages";
import { FolderPicker, type FolderCrumb } from "../FolderPicker";
import { Button, Dialog, TextField } from "../ui";

export function CreateFolderDialog({
  open,
  messages,
  pending,
  onClose,
  onSubmit,
}: {
  open: boolean;
  messages: Messages;
  pending: boolean;
  onClose: () => void;
  onSubmit: (name: string) => void;
}) {
  const [name, setName] = useState("");
  return (
    <Dialog
      open={open}
      title={messages.newFolder}
      onClose={onClose}
      footer={
        <>
          <Button onClick={onClose}>{messages.cancel}</Button>
          <Button
            variant="primary"
            disabled={!name.trim() || pending}
            onClick={() => onSubmit(name.trim())}
          >
            {messages.create}
          </Button>
        </>
      }
    >
      <TextField label={messages.folderName} value={name} onChange={setName} />
    </Dialog>
  );
}

export function RenameDialog({
  item,
  open,
  messages,
  pending,
  onClose,
  onSubmit,
}: {
  item?: SharePointItem;
  open: boolean;
  messages: Messages;
  pending: boolean;
  onClose: () => void;
  onSubmit: (name: string) => void;
}) {
  const [name, setName] = useState(item?.name ?? "");
  useEffect(() => {
    setName(item?.name ?? "");
  }, [item?.name]);
  return (
    <Dialog
      open={open}
      title={messages.rename}
      onClose={onClose}
      footer={
        <>
          <Button onClick={onClose}>{messages.cancel}</Button>
          <Button variant="primary" disabled={!name.trim() || pending} onClick={() => onSubmit(name.trim())}>
            {messages.save}
          </Button>
        </>
      }
    >
      <TextField label={messages.newName} value={name} onChange={setName} />
    </Dialog>
  );
}

export function DeleteDialog({
  items,
  open,
  messages,
  pending,
  onClose,
  onConfirm,
}: {
  items: SharePointItem[];
  open: boolean;
  messages: Messages;
  pending: boolean;
  onClose: () => void;
  onConfirm: () => void;
}) {
  const names = items.map((item) => item.name).filter(Boolean);

  return (
    <Dialog
      open={open}
      title={messages.delete}
      onClose={onClose}
      footer={
        <>
          <Button onClick={onClose}>{messages.cancel}</Button>
          <Button variant="danger" disabled={pending || names.length === 0} onClick={onConfirm}>
            {messages.delete}
          </Button>
        </>
      }
    >
      <p className="spm-text-sm">{names.length > 1 ? messages.confirmDeleteMany : messages.confirmDelete}</p>
      {names.length === 1 ? <p className="spm-mt-2 spm-font-medium">{names[0]}</p> : null}
      {names.length > 1 ? <p className="spm-mt-2 spm-text-sm spm-font-medium">{names.length} {messages.selected}</p> : null}
    </Dialog>
  );
}

export function CopyMoveDialog({
  mode,
  open,
  messages,
  rootId,
  rootName,
  initialCrumbs,
  excludeIds,
  singleItem,
  pending,
  copyProgress,
  showCopyProgress,
  onClose,
  onSubmit,
}: {
  mode: "copy" | "move";
  open: boolean;
  messages: Messages;
  rootId: string;
  rootName: string;
  initialCrumbs: FolderCrumb[];
  excludeIds: string[];
  singleItem?: SharePointItem;
  pending: boolean;
  copyProgress?: number;
  showCopyProgress?: boolean;
  onClose: () => void;
  onSubmit: (destinationParentId: string, newName?: string) => void;
}) {
  const [destination, setDestination] = useState(rootId);
  const [newName, setNewName] = useState(singleItem?.name ?? "");
  const blocked = excludeIds.includes(destination);

  useEffect(() => {
    setNewName(singleItem?.name ?? "");
  }, [singleItem?.name, open]);

  return (
    <Dialog
      open={open}
      title={mode === "copy" ? messages.copy : messages.move}
      onClose={onClose}
      footer={
        <>
          <Button onClick={onClose}>{messages.cancel}</Button>
          <Button
            variant="primary"
            disabled={!destination || pending || blocked}
            onClick={() => onSubmit(destination, newName.trim() || undefined)}
          >
            {mode === "copy" ? messages.copy : messages.move}
          </Button>
        </>
      }
    >
      {singleItem ? (
        <div className="spm-mb-3">
          <TextField label={messages.newName} value={newName} onChange={setNewName} />
        </div>
      ) : null}
      <FolderPicker
        open={open}
        rootId={rootId}
        rootName={rootName}
        initialCrumbs={initialCrumbs}
        excludeIds={excludeIds}
        messages={messages}
        onFolderChange={setDestination}
      />
      {blocked ? <p className="spm-mt-2 spm-text-xs spm-text-sp-danger">{messages.invalidDestination}</p> : null}
      {showCopyProgress && mode === "copy" && pending ? (
        <div className="spm-mt-3">
          <div className="spm-mb-1 spm-text-xs spm-text-sp-muted">
            {messages.copyInProgress} {copyProgress !== undefined ? `${copyProgress}%` : ""}
          </div>
          <div className="spm-h-2 spm-w-full spm-overflow-hidden spm-rounded-full spm-bg-slate-200">
            <div
              className="spm-h-full spm-bg-sp-primary spm-transition-all"
              style={{ width: `${copyProgress ?? 0}%` }}
            />
          </div>
        </div>
      ) : null}
    </Dialog>
  );
}

export function CheckinDialog({
  item,
  open,
  messages,
  pending,
  onClose,
  onSubmit,
}: {
  item?: SharePointItem;
  open: boolean;
  messages: Messages;
  pending: boolean;
  onClose: () => void;
  onSubmit: (comment?: string) => void;
}) {
  const [comment, setComment] = useState("");

  useEffect(() => {
    if (!open) setComment("");
  }, [open]);

  return (
    <Dialog
      open={open}
      title={`${messages.checkin}${item ? `: ${item.name}` : ""}`}
      onClose={onClose}
      footer={
        <>
          <Button onClick={onClose}>{messages.cancel}</Button>
          <Button variant="primary" disabled={pending} onClick={() => onSubmit(comment.trim() || undefined)}>
            {messages.checkin}
          </Button>
        </>
      }
    >
      <TextField label={messages.checkinComment} value={comment} onChange={setComment} />
    </Dialog>
  );
}
