import { useEffect, useState } from "react";
import {
  toInviteRecipient,
  type ConflictBehavior,
  type DirectoryPerson,
  type InviteRecipient,
  type SharePointItem,
} from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../../i18n/messages";
import { Checkbox, Field, Input, ProgressBar, Select, Textarea } from "@fluentui/react-components";
import { PeoplePicker } from "../PeoplePicker";
import { Button, Dialog } from "../ui";

export function UploadDialog({
  open,
  messages,
  progress,
  pending,
  onClose,
  onUpload,
  onCancel,
}: {
  open: boolean;
  messages: Messages;
  progress?: number;
  pending: boolean;
  onClose: () => void;
  onUpload: (files: FileList, conflictBehavior: ConflictBehavior) => void;
  onCancel?: () => void;
}) {
  const [conflictBehavior, setConflictBehavior] = useState<ConflictBehavior>("rename");

  useEffect(() => {
    if (!open) setConflictBehavior("rename");
  }, [open]);

  return (
    <Dialog
      open={open}
      title={messages.upload}
      onClose={onClose}
      footer={
        <>
          {pending && onCancel ? <Button onClick={onCancel}>{messages.cancelUpload}</Button> : null}
          <Button onClick={onClose}>{messages.cancel}</Button>
        </>
      }
    >
      <div className="spm-space-y-4">
        <Field label={messages.uploadConflict}>
          <Select
            value={conflictBehavior}
            disabled={pending}
            onChange={(event) => setConflictBehavior(event.target.value as ConflictBehavior)}
          >
            <option value="rename">{messages.conflictRename}</option>
            <option value="replace">{messages.conflictReplace}</option>
            <option value="fail">{messages.conflictFail}</option>
          </Select>
        </Field>
        <label className="spm-upload-dropzone">
          <span>{messages.uploadFiles}</span>
          <input
            type="file"
            multiple
            disabled={pending}
            onChange={(event) => {
              if (event.target.files?.length) onUpload(event.target.files, conflictBehavior);
            }}
          />
        </label>
      </div>
      {pending ? (
        <div className="spm-mt-4 spm-space-y-2">
          <div className="spm-flex spm-justify-between spm-text-sm">
            <span>{messages.uploading}</span><span>{Math.round(progress ?? 0)}%</span>
          </div>
          <ProgressBar value={(progress ?? 0) / 100} />
        </div>
      ) : null}
    </Dialog>
  );
}

export function ShareDialog({
  item,
  open,
  messages,
  pending,
  error,
  createdLinkUrl,
  onClose,
  onInvite,
  onCreateLink,
}: {
  item?: SharePointItem;
  open: boolean;
  messages: Messages;
  pending: boolean;
  error?: string;
  createdLinkUrl?: string;
  onClose: () => void;
  onInvite: (recipients: InviteRecipient[], role: "read" | "write", message: string, notify: boolean) => void;
  onCreateLink: (scope: "anonymous" | "organization", type: "view" | "edit", expiration?: string) => void;
}) {
  const [selected, setSelected] = useState<DirectoryPerson[]>([]);
  const [role, setRole] = useState<"read" | "write">("read");
  const [message, setMessage] = useState("");
  const [notify, setNotify] = useState(true);
  const [scope, setScope] = useState<"organization" | "anonymous" | "users">("organization");
  const [linkType, setLinkType] = useState<"view" | "edit">("view");
  const [expiration, setExpiration] = useState("");

  useEffect(() => {
    if (!open) {
      setSelected([]);
      setRole("read");
      setMessage("");
      setNotify(true);
      setScope("organization");
    }
  }, [open]);

  const recipients = selected
    .map(toInviteRecipient)
    .filter((recipient): recipient is InviteRecipient => Boolean(recipient));

  function handleCreateLink() {
    if (scope === "users") {
      if (recipients.length === 0) return;
      onInvite(recipients, role, message, notify);
      return;
    }
    onCreateLink(scope, linkType, expiration ? new Date(expiration).toISOString() : undefined);
  }

  return (
    <Dialog
      open={open}
      title={`${messages.share}${item ? `: ${item.name}` : ""}`}
      onClose={onClose}
      footer={
        <>
          <Button
            disabled={pending || (scope === "users" && recipients.length === 0)}
            onClick={handleCreateLink}
          >
            {scope === "users" ? messages.send : messages.copyLink}
          </Button>
          {scope !== "users" ? (
            <Button
              variant="primary"
              disabled={recipients.length === 0 || pending}
              onClick={() => onInvite(recipients, role, message, notify)}
            >
              {messages.send}
            </Button>
          ) : null}
        </>
      }
    >
      <div className="spm-space-y-4">
        {error ? <p className="spm-text-sm spm-text-sp-danger">{error}</p> : null}

        <PeoplePicker open={open} selected={selected} messages={messages} onChange={setSelected} />

        <div className="spm-flex spm-flex-wrap spm-items-center spm-gap-2">
          <Select
            value={role}
            onChange={(event) => setRole(event.target.value as "read" | "write")}
          >
            <option value="read">{messages.canView}</option>
            <option value="write">{messages.canEdit}</option>
          </Select>
          <Checkbox label={messages.notify} checked={notify} onChange={(_, data) => setNotify(Boolean(data.checked))} />
        </div>

        <Textarea
          resize="vertical"
          placeholder={messages.message}
          value={message}
          onChange={(_, data) => setMessage(data.value)}
        />

        <hr />
        <div className="spm-space-y-2">
          <Select
            value={scope}
            onChange={(event) => setScope(event.target.value as typeof scope)}
          >
            <option value="organization">{messages.organization}</option>
            <option value="users">{messages.specificPeople}</option>
            <option value="anonymous">{messages.anyone}</option>
          </Select>
          {scope !== "users" ? (
            <>
              <Select
                value={linkType}
                onChange={(event) => setLinkType(event.target.value as "view" | "edit")}
              >
                <option value="view">{messages.view}</option>
                <option value="edit">{messages.edit}</option>
              </Select>
              <Input
                type="datetime-local"
                value={expiration}
                onChange={(_, data) => setExpiration(data.value)}
              />
              {createdLinkUrl ? (
                <p className="spm-break-all spm-text-xs spm-text-sp-muted">
                  {messages.copyLinkSuccess}: {createdLinkUrl}
                </p>
              ) : null}
            </>
          ) : (
            <p className="spm-text-xs spm-text-sp-muted">{messages.peopleSearchPlaceholder}</p>
          )}
        </div>
      </div>
    </Dialog>
  );
}
