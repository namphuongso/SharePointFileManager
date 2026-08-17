import { useEffect, useState } from "react";
import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../../i18n/messages";
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
  item,
  open,
  messages,
  pending,
  onClose,
  onConfirm,
}: {
  item?: SharePointItem;
  open: boolean;
  messages: Messages;
  pending: boolean;
  onClose: () => void;
  onConfirm: () => void;
}) {
  return (
    <Dialog
      open={open}
      title={messages.delete}
      onClose={onClose}
      footer={
        <>
          <Button onClick={onClose}>{messages.cancel}</Button>
          <Button variant="danger" disabled={pending} onClick={onConfirm}>
            {messages.delete}
          </Button>
        </>
      }
    >
      <p className="spm-text-sm">{messages.confirmDelete}</p>
      {item ? <p className="spm-mt-2 spm-font-medium">{item.name}</p> : null}
    </Dialog>
  );
}

export function CopyMoveDialog({
  mode,
  open,
  messages,
  currentFolderId,
  pending,
  onClose,
  onSubmit,
}: {
  mode: "copy" | "move";
  open: boolean;
  messages: Messages;
  currentFolderId: string;
  pending: boolean;
  onClose: () => void;
  onSubmit: (destinationParentId: string) => void;
}) {
  const [destination, setDestination] = useState(currentFolderId);
  return (
    <Dialog
      open={open}
      title={mode === "copy" ? messages.copy : messages.move}
      onClose={onClose}
      footer={
        <>
          <Button onClick={onClose}>{messages.cancel}</Button>
          <Button variant="primary" disabled={!destination || pending} onClick={() => onSubmit(destination)}>
            {mode === "copy" ? messages.copy : messages.move}
          </Button>
        </>
      }
    >
      <TextField label={messages.destination} value={destination} onChange={setDestination} />
      <p className="spm-mt-2 spm-text-xs spm-text-sp-muted">SharePoint item id of the destination folder.</p>
    </Dialog>
  );
}
