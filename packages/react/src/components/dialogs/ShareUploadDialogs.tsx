import { useState } from "react";
import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../../i18n/messages";
import { Button, Dialog } from "../ui";

export function UploadDialog({
  open,
  messages,
  progress,
  pending,
  onClose,
  onUpload,
}: {
  open: boolean;
  messages: Messages;
  progress?: number;
  pending: boolean;
  onClose: () => void;
  onUpload: (files: FileList) => void;
}) {
  return (
    <Dialog
      open={open}
      title={messages.upload}
      onClose={onClose}
      footer={<Button onClick={onClose}>{messages.cancel}</Button>}
    >
      <input
        type="file"
        multiple
        disabled={pending}
        onChange={(event) => {
          if (event.target.files?.length) onUpload(event.target.files);
        }}
      />
      {pending ? (
        <div className="spm-mt-3">
          <div className="spm-text-sm">{messages.uploading}</div>
          <div className="spm-mt-1 spm-h-2 spm-w-full spm-rounded spm-bg-slate-100">
            <div className="spm-h-2 spm-rounded spm-bg-sp-primary" style={{ width: `${progress ?? 0}%` }} />
          </div>
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
  onClose,
  onInvite,
  onCreateLink,
}: {
  item?: SharePointItem;
  open: boolean;
  messages: Messages;
  pending: boolean;
  error?: string;
  onClose: () => void;
  onInvite: (email: string, role: "read" | "write", message: string, notify: boolean) => void;
  onCreateLink: (scope: "anonymous" | "organization" | "users", type: "view" | "edit", expiration?: string) => void;
}) {
  const [email, setEmail] = useState("");
  const [role, setRole] = useState<"read" | "write">("read");
  const [message, setMessage] = useState("");
  const [notify, setNotify] = useState(true);
  const [scope, setScope] = useState<"organization" | "anonymous" | "users">("organization");
  const [linkType, setLinkType] = useState<"view" | "edit">("view");
  const [expiration, setExpiration] = useState("");

  return (
    <Dialog
      open={open}
      title={`${messages.share}${item ? `: ${item.name}` : ""}`}
      onClose={onClose}
    >
      <div className="spm-space-y-4">
        {error ? <p className="spm-text-sm spm-text-sp-danger">{error}</p> : null}
        <div className="spm-space-y-2">
          <input
            className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-3 spm-py-2"
            placeholder={messages.email}
            value={email}
            onChange={(event) => setEmail(event.target.value)}
          />
          <select
            className="spm-rounded-md spm-border spm-border-sp-border spm-px-3 spm-py-2"
            value={role}
            onChange={(event) => setRole(event.target.value as "read" | "write")}
          >
            <option value="read">{messages.view}</option>
            <option value="write">{messages.edit}</option>
          </select>
          <textarea
            className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-3 spm-py-2"
            placeholder={messages.message}
            value={message}
            onChange={(event) => setMessage(event.target.value)}
          />
          <label className="spm-flex spm-items-center spm-gap-2 spm-text-sm">
            <input type="checkbox" checked={notify} onChange={(event) => setNotify(event.target.checked)} />
            {messages.notify}
          </label>
          <Button
            variant="primary"
            disabled={!email || pending}
            onClick={() => onInvite(email, role, message, notify)}
          >
            {messages.grantAccess}
          </Button>
        </div>
        <hr />
        <div className="spm-space-y-2">
          <select
            className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-3 spm-py-2"
            value={scope}
            onChange={(event) => setScope(event.target.value as typeof scope)}
          >
            <option value="organization">{messages.organization}</option>
            <option value="users">{messages.specificPeople}</option>
            <option value="anonymous">{messages.anyone}</option>
          </select>
          <select
            className="spm-rounded-md spm-border spm-border-sp-border spm-px-3 spm-py-2"
            value={linkType}
            onChange={(event) => setLinkType(event.target.value as "view" | "edit")}
          >
            <option value="view">{messages.view}</option>
            <option value="edit">{messages.edit}</option>
          </select>
          <input
            type="datetime-local"
            className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-3 spm-py-2"
            value={expiration}
            onChange={(event) => setExpiration(event.target.value)}
          />
          <Button
            disabled={pending}
            onClick={() =>
              onCreateLink(scope, linkType, expiration ? new Date(expiration).toISOString() : undefined)
            }
          >
            {messages.createLink}
          </Button>
        </div>
      </div>
    </Dialog>
  );
}
