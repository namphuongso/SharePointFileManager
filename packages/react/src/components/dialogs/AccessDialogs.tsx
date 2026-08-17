import type { ReactNode } from "react";
import type { FileVersion, PreviewInfo, SharePointItem, SharePointPermission } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../../i18n/messages";
import { Button, Dialog, formatBytes, formatDate } from "../ui";

export function ManageAccessDialog({
  item,
  open,
  messages,
  locale,
  permissions,
  loading,
  error,
  onClose,
  onRemove,
  onCopyLink,
}: {
  item?: SharePointItem;
  open: boolean;
  messages: Messages;
  locale: string;
  permissions: SharePointPermission[];
  loading: boolean;
  error?: string;
  onClose: () => void;
  onRemove: (permissionId: string) => void;
  onCopyLink: (url: string) => void;
}) {
  const people = permissions.filter((permission) => permission.kind === "user");
  const groups = permissions.filter((permission) => permission.kind === "group" || permission.kind === "siteGroup");
  const links = permissions.filter((permission) => permission.kind === "link");

  return (
    <Dialog open={open} title={`${messages.manageAccess}${item ? `: ${item.name}` : ""}`} onClose={onClose}>
      {loading ? <p className="spm-text-sm">{messages.loading}</p> : null}
      {error ? <p className="spm-text-sm spm-text-sp-danger">{error}</p> : null}
      <Section title={messages.people}>
        {people.map((permission) => (
          <PermissionRow
            key={permission.id}
            title={permission.grantedTo?.displayName ?? permission.grantedTo?.email ?? permission.id}
            subtitle={permission.roles.join(", ")}
            inherited={permission.inherited}
            messages={messages}
            canRemove={permission.canRemove}
            onRemove={() => onRemove(permission.id)}
          />
        ))}
      </Section>
      <Section title={messages.groups}>
        {groups.map((permission) => (
          <PermissionRow
            key={permission.id}
            title={permission.grantedToGroup?.displayName ?? permission.id}
            subtitle={permission.roles.join(", ")}
            inherited={permission.inherited}
            messages={messages}
            canRemove={permission.canRemove}
            onRemove={() => onRemove(permission.id)}
          />
        ))}
      </Section>
      <Section title={messages.links}>
        {links.map((permission) => (
          <div key={permission.id} className="spm-flex spm-items-center spm-justify-between spm-gap-2 spm-py-2">
            <div>
              <div className="spm-text-sm">{permission.link?.scope} · {permission.link?.type}</div>
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

function PermissionRow({
  title,
  subtitle,
  inherited,
  canRemove,
  messages,
  onRemove,
}: {
  title: string;
  subtitle: string;
  inherited: boolean;
  canRemove: boolean;
  messages: Messages;
  onRemove: () => void;
}) {
  return (
    <div className="spm-flex spm-items-center spm-justify-between spm-py-2">
      <div>
        <div className="spm-text-sm">{title}</div>
        <div className="spm-text-xs spm-text-sp-muted">
          {subtitle}
          {inherited ? ` · ${messages.inherited}` : ""}
        </div>
      </div>
      {canRemove ? (
        <Button variant="danger" onClick={onRemove}>
          {messages.revoke}
        </Button>
      ) : null}
    </div>
  );
}

export function PreviewDialog({
  item,
  open,
  preview,
  messages,
  onClose,
}: {
  item?: SharePointItem;
  open: boolean;
  preview?: PreviewInfo;
  messages: Messages;
  onClose: () => void;
}) {
  return (
    <Dialog open={open} title={`${messages.preview}${item ? `: ${item.name}` : ""}`} onClose={onClose}>
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
  onClose,
  onRestore,
}: {
  item?: SharePointItem;
  open: boolean;
  versions: FileVersion[];
  locale: string;
  messages: Messages;
  onClose: () => void;
  onRestore: (versionId: string) => void;
}) {
  return (
    <Dialog open={open} title={messages.versionHistory} onClose={onClose}>
      {versions.map((version) => (
        <div key={version.id} className="spm-flex spm-items-center spm-justify-between spm-py-2">
          <div>
            <div className="spm-text-sm">{formatDate(version.lastModifiedDateTime, locale)}</div>
            <div className="spm-text-xs spm-text-sp-muted">
              {version.lastModifiedBy?.displayName} · {formatBytes(version.size)}
            </div>
          </div>
          <Button onClick={() => onRestore(version.id)}>{messages.restore}</Button>
        </div>
      ))}
      {item && versions.length === 0 ? <p className="spm-text-sm">{messages.empty}</p> : null}
    </Dialog>
  );
}
