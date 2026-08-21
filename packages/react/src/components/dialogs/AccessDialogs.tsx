import { useEffect, useState } from "react";
import type { ReactNode } from "react";
import {
  toInviteRecipient,
  type DirectoryPerson,
  type FileVersion,
  type InviteRecipient,
  type ListColumn,
  type ListItemFields,
  type PreviewInfo,
  type SharePointItem,
  type SharePointPermission,
} from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../../i18n/messages";
import { PeoplePicker } from "../PeoplePicker";
import { Button, Dialog, formatBytes, formatDate } from "../ui";

export function ManageAccessDialog({
  item,
  open,
  messages,
  locale,
  permissions,
  loading,
  error,
  grantPending,
  grantSuccess,
  onClose,
  onRemove,
  onCopyLink,
  onChangeRole,
  onGrant,
}: {
  item?: SharePointItem;
  open: boolean;
  messages: Messages;
  locale: string;
  permissions: SharePointPermission[];
  loading: boolean;
  error?: string;
  grantPending?: boolean;
  grantSuccess?: boolean;
  onClose: () => void;
  onRemove: (permissionId: string) => void;
  onCopyLink: (url: string) => void;
  onChangeRole?: (permissionId: string, roles: string[]) => void;
  onGrant?: (recipients: InviteRecipient[], role: "read" | "write") => void;
}) {
  const [selected, setSelected] = useState<DirectoryPerson[]>([]);
  const [grantRole, setGrantRole] = useState<"read" | "write">("read");

  const people = permissions.filter((permission) => permission.kind === "user");
  const groups = permissions.filter((permission) => permission.kind === "group" || permission.kind === "siteGroup");
  const links = permissions.filter((permission) => permission.kind === "link");

  const recipients = selected
    .map(toInviteRecipient)
    .filter((recipient): recipient is InviteRecipient => Boolean(recipient));

  return (
    <Dialog open={open} title={`${messages.manageAccess}${item ? `: ${item.name}` : ""}`} onClose={onClose}>
      {loading ? <p className="spm-text-sm">{messages.loading}</p> : null}
      {error ? <p className="spm-text-sm spm-text-sp-danger">{error}</p> : null}

      {onGrant ? (
        <section className="spm-mb-4 spm-rounded-md spm-border spm-border-sp-border spm-p-3">
          <h3 className="spm-mb-2 spm-text-xs spm-font-semibold spm-uppercase spm-text-sp-muted">
            {messages.grantAccess}
          </h3>
          <PeoplePicker open={open} selected={selected} messages={messages} onChange={setSelected} />
          <div className="spm-mt-2 spm-flex spm-flex-wrap spm-items-center spm-gap-2">
            <select
              className="spm-rounded-md spm-border spm-border-sp-border spm-px-2 spm-py-1 spm-text-sm"
              value={grantRole}
              onChange={(event) => setGrantRole(event.target.value as "read" | "write")}
            >
              <option value="read">{messages.canView}</option>
              <option value="write">{messages.canEdit}</option>
            </select>
            <Button
              variant="primary"
              disabled={recipients.length === 0 || grantPending}
              onClick={() => {
                onGrant(recipients, grantRole);
                setSelected([]);
              }}
            >
              {messages.grantAccess}
            </Button>
          </div>
          {grantSuccess ? <p className="spm-mt-2 spm-text-xs spm-text-green-700">{messages.grantAccessSuccess}</p> : null}
        </section>
      ) : null}

      <Section title={messages.people}>
        {people.map((permission) => (
          <PermissionRow
            key={permission.id}
            title={permission.grantedTo?.displayName ?? permission.grantedTo?.email ?? permission.id}
            roles={permission.roles}
            inherited={permission.inherited}
            messages={messages}
            canRemove={permission.canRemove}
            onRemove={() => onRemove(permission.id)}
            // Keep user/group ACL entries read-only to avoid mutating inherited or ambiguous permissions.
            onChangeRole={undefined}
          />
        ))}
      </Section>
      <Section title={messages.groups}>
        {groups.map((permission) => (
          <PermissionRow
            key={permission.id}
            title={permission.grantedToGroup?.displayName ?? permission.id}
            roles={permission.roles}
            inherited={permission.inherited}
            messages={messages}
            canRemove={permission.canRemove}
            onRemove={() => onRemove(permission.id)}
            // Keep user/group ACL entries read-only to avoid mutating inherited or ambiguous permissions.
            onChangeRole={undefined}
          />
        ))}
      </Section>
      <Section title={messages.links}>
        {links.map((permission) => (
          <LinkPermissionRow
            key={permission.id}
            permission={permission}
            locale={locale}
            messages={messages}
            onCopyLink={onCopyLink}
            onRemove={onRemove}
          />
        ))}
      </Section>
    </Dialog>
  );
}

function Section({ title, children }: { title: string; children: ReactNode }) {
  return (
    <section className="spm-mb-4">
      <h3 className="spm-mb-1 spm-text-xs spm-font-semibold spm-uppercase spm-text-sp-muted">{title}</h3>
      <div>{children}</div>
    </section>
  );
}

function LinkPermissionRow({
  permission,
  locale,
  messages,
  onCopyLink,
  onRemove,
}: {
  permission: SharePointPermission;
  locale: string;
  messages: Messages;
  onCopyLink: (url: string) => void;
  onRemove: (permissionId: string) => void;
}) {
  return (
    <div className="spm-flex spm-flex-wrap spm-items-center spm-justify-between spm-gap-2 spm-py-2">
      <div>
        <div className="spm-text-sm">
          {permission.link?.scope} · {permission.link?.type}
        </div>
        {permission.expirationDateTime ? (
          <div className="spm-text-xs spm-text-sp-muted">
            {messages.expiration}: {formatDate(permission.expirationDateTime, locale)}
          </div>
        ) : null}
      </div>
      <div className="spm-flex spm-gap-2">
        {permission.link?.webUrl ? (
          <Button onClick={() => onCopyLink(permission.link!.webUrl!)}>{messages.copyLink}</Button>
        ) : null}
        {permission.canRemove ? (
          <Button variant="danger" onClick={() => onRemove(permission.id)}>
            {messages.revoke}
          </Button>
        ) : null}
      </div>
    </div>
  );
}

function PermissionRow({
  title,
  roles,
  inherited,
  canRemove,
  messages,
  onRemove,
  onChangeRole,
}: {
  title: string;
  roles: string[];
  inherited: boolean;
  canRemove: boolean;
  messages: Messages;
  onRemove: () => void;
  onChangeRole?: (role: "read" | "write") => void;
}) {
  const currentRole = roles.includes("write") || roles.includes("edit") ? "write" : "read";

  return (
    <div className="spm-flex spm-items-center spm-justify-between spm-gap-2 spm-py-2">
      <div className="spm-min-w-0">
        <div className="spm-truncate spm-text-sm">{title}</div>
        <div className="spm-text-xs spm-text-sp-muted">{inherited ? messages.inherited : ""}</div>
      </div>
      <div className="spm-flex spm-shrink-0 spm-items-center spm-gap-2">
        {onChangeRole ? (
          <select
            className="spm-rounded-md spm-border spm-border-sp-border spm-px-2 spm-py-1 spm-text-xs"
            value={currentRole}
            onChange={(event) => onChangeRole(event.target.value as "read" | "write")}
          >
            <option value="read">{messages.canView}</option>
            <option value="write">{messages.canEdit}</option>
          </select>
        ) : (
          <span className="spm-text-xs spm-text-sp-muted">
            {currentRole === "write" ? messages.canEdit : messages.canView}
          </span>
        )}
        {canRemove ? (
          <Button variant="danger" onClick={onRemove}>
            {messages.revoke}
          </Button>
        ) : null}
      </div>
    </div>
  );
}

export function PropertiesDialog({
  item,
  open,
  messages,
  locale,
  loading,
  listColumns,
  listItemFields,
  metadataPending,
  onClose,
  onSaveMetadata,
}: {
  item?: SharePointItem;
  open: boolean;
  messages: Messages;
  locale: string;
  loading?: boolean;
  listColumns?: ListColumn[];
  listItemFields?: ListItemFields;
  metadataPending?: boolean;
  onClose: () => void;
  onSaveMetadata?: (fields: Record<string, string | number | boolean | null>) => void;
}) {
  const [draft, setDraft] = useState<Record<string, string>>({});

  useEffect(() => {
    if (!open) {
      setDraft({});
      return;
    }
    const next: Record<string, string> = {};
    for (const column of listColumns ?? []) {
      const value = listItemFields?.fields[column.name] ?? item?.metadata?.[column.name];
      if (value === null || value === undefined) {
        next[column.name] = "";
        continue;
      }
      if (column.type === "dateTime" && typeof value === "string") {
        next[column.name] = value.slice(0, 10);
        continue;
      }
      next[column.name] = String(value);
    }
    setDraft(next);
  }, [open, item?.id, listColumns, listItemFields, item?.metadata]);

  const editableColumns = (listColumns ?? []).filter((column) => !column.readOnly);

  return (
    <Dialog open={open} title={`${messages.properties}${item ? `: ${item.name}` : ""}`} onClose={onClose}>
      {loading || !item ? <p className="spm-text-sm">{messages.loading}</p> : null}
      {item && !loading ? (
        <dl className="spm-space-y-2 spm-text-sm">
          <PropertyRow label={messages.name} value={item.name} />
          <PropertyRow
            label={messages.itemType}
            value={item.type === "folder" ? messages.folder : messages.file}
          />
          <PropertyRow label={messages.size} value={item.type === "file" ? formatBytes(item.size) : "—"} />
          <PropertyRow label={messages.modified} value={formatDate(item.lastModifiedDateTime, locale)} />
          <PropertyRow label={messages.createdBy} value={item.createdBy?.displayName ?? item.createdBy?.email ?? "—"} />
          <PropertyRow
            label={messages.modifiedBy}
            value={item.lastModifiedBy?.displayName ?? item.lastModifiedBy?.email ?? "—"}
          />
          {item.mimeType ? <PropertyRow label={messages.type} value={item.mimeType} /> : null}
          {item.contentType || listItemFields?.contentType ? (
            <PropertyRow label={messages.contentType} value={listItemFields?.contentType ?? item.contentType ?? "—"} />
          ) : null}
          {item.sensitivityLabel ? (
            <PropertyRow label={messages.sensitivityLabel} value={item.sensitivityLabel} />
          ) : null}
          {(listColumns ?? []).length > 0 ? (
            <section className="spm-pt-2">
              <h3 className="spm-mb-2 spm-text-xs spm-font-semibold spm-uppercase spm-text-sp-muted">{messages.metadata}</h3>
              {(listColumns ?? []).map((column) => {
                const readOnly = column.readOnly || !onSaveMetadata;
                const value = draft[column.name] ?? "";
                if (readOnly) {
                  return <PropertyRow key={column.id} label={column.displayName} value={value || "—"} />;
                }
                return (
                  <label key={column.id} className="spm-mb-2 spm-block spm-text-sm">
                    <span className="spm-mb-1 spm-block spm-text-sp-muted">{column.displayName}</span>
                    {renderMetadataInput(column, value, (next) =>
                      setDraft((current) => ({ ...current, [column.name]: next }))
                    )}
                  </label>
                );
              })}
              {editableColumns.length > 0 && onSaveMetadata ? (
                <Button
                  variant="primary"
                  disabled={metadataPending}
                  onClick={() => {
                    const payload: Record<string, string | number | boolean | null> = {};
                    for (const column of editableColumns) {
                      const value = draft[column.name];
                      payload[column.name] = parseDraftValue(column, value ?? "");
                    }
                    onSaveMetadata(payload);
                  }}
                >
                  {messages.save}
                </Button>
              ) : null}
            </section>
          ) : null}
        </dl>
      ) : null}
    </Dialog>
  );
}

function PropertyRow({ label, value }: { label: string; value: string }) {
  return (
    <div className="spm-flex spm-gap-2">
      <dt className="spm-w-28 spm-shrink-0 spm-text-sp-muted">{label}</dt>
      <dd className="spm-min-w-0 spm-break-all">{value}</dd>
    </div>
  );
}

function parseDraftValue(column: ListColumn, value: string): string | number | boolean | null {
  const trimmed = value.trim();
  if (!trimmed) return null;
  if (column.type === "number") {
    const parsed = Number(trimmed);
    return Number.isFinite(parsed) ? parsed : null;
  }
  if (column.type === "boolean") {
    if (trimmed.toLowerCase() === "true") return true;
    if (trimmed.toLowerCase() === "false") return false;
  }
  return trimmed;
}

function renderMetadataInput(
  column: ListColumn,
  value: string,
  onChange: (next: string) => void,
) {
  if (column.type === "boolean") {
    return (
      <select
        className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-2 spm-py-1.5"
        value={value}
        onChange={(event) => onChange(event.target.value)}
      >
        <option value="">—</option>
        <option value="true">true</option>
        <option value="false">false</option>
      </select>
    );
  }
  if (column.type === "dateTime") {
    return (
      <input
        type="date"
        className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-2 spm-py-1.5"
        value={value}
        onChange={(event) => onChange(event.target.value)}
      />
    );
  }
  if (column.type === "number") {
    return (
      <input
        type="number"
        className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-2 spm-py-1.5"
        value={value}
        onChange={(event) => onChange(event.target.value)}
      />
    );
  }
  return (
    <input
      className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-2 spm-py-1.5"
      value={value}
      onChange={(event) => onChange(event.target.value)}
    />
  );
}

export function PreviewDialog({
  item,
  open,
  preview,
  messages,
  onClose,
  onPrint,
}: {
  item?: SharePointItem;
  open: boolean;
  preview?: PreviewInfo;
  messages: Messages;
  onClose: () => void;
  onPrint?: () => void;
}) {
  return (
    <Dialog
      open={open}
      title={`${messages.preview}${item ? `: ${item.name}` : ""}`}
      onClose={onClose}
      footer={onPrint && preview?.getUrl ? <Button onClick={onPrint}>{messages.print}</Button> : undefined}
    >
      {preview?.getUrl ? (
        <iframe title={item?.name} src={preview.getUrl} className="spm-h-[60vh] spm-w-full spm-border-0" />
      ) : (
        <p className="spm-text-sm spm-text-sp-muted">{messages.loading}</p>
      )}
    </Dialog>
  );
}

export function VersionHistoryDialog({
  item,
  open,
  versions,
  locale,
  messages,
  loading,
  error,
  restorePending,
  restoreSuccess,
  downloadPending,
  onClose,
  onRestore,
  onDownload,
}: {
  item?: SharePointItem;
  open: boolean;
  versions: FileVersion[];
  locale: string;
  messages: Messages;
  loading?: boolean;
  error?: string;
  restorePending?: boolean;
  restoreSuccess?: boolean;
  downloadPending?: boolean;
  onClose: () => void;
  onRestore: (versionId: string) => void;
  onDownload: (versionId: string) => void;
}) {
  return (
    <Dialog open={open} title={`${messages.versionHistory}${item ? `: ${item.name}` : ""}`} onClose={onClose}>
      {loading ? <p className="spm-text-sm">{messages.loading}</p> : null}
      {error ? <p className="spm-text-sm spm-text-sp-danger">{error}</p> : null}
      {restoreSuccess ? <p className="spm-mb-2 spm-text-sm spm-text-green-700">{messages.restoreSuccess}</p> : null}
      {versions.map((version) => (
        <div key={version.id} className="spm-flex spm-items-center spm-justify-between spm-gap-2 spm-py-2">
          <div>
            <div className="spm-text-sm">{formatDate(version.lastModifiedDateTime, locale)}</div>
            <div className="spm-text-xs spm-text-sp-muted">
              {version.lastModifiedBy?.displayName} · {formatBytes(version.size)}
            </div>
          </div>
          <div className="spm-flex spm-gap-2">
            <Button disabled={downloadPending} onClick={() => onDownload(version.id)}>
              {messages.downloadVersion}
            </Button>
            <Button disabled={restorePending} onClick={() => onRestore(version.id)}>
              {messages.restore}
            </Button>
          </div>
        </div>
      ))}
      {item && !loading && versions.length === 0 ? <p className="spm-text-sm">{messages.empty}</p> : null}
    </Dialog>
  );
}
