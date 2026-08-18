import type { DriveItemActivity } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../i18n/messages";
import { Button, Dialog, formatDate } from "./ui";

export function ActivityDialog({
  open,
  title,
  activities,
  locale,
  messages,
  loading,
  error,
  onClose,
}: {
  open: boolean;
  title?: string;
  activities: DriveItemActivity[];
  locale: string;
  messages: Messages;
  loading?: boolean;
  error?: string;
  onClose: () => void;
}) {
  return (
    <Dialog open={open} title={title ?? messages.activityLog} onClose={onClose}>
      {loading ? <p className="spm-text-sm">{messages.loading}</p> : null}
      {error ? <p className="spm-text-sm spm-text-sp-danger">{error}</p> : null}
      {!loading && !error && activities.length === 0 ? <p className="spm-text-sm spm-text-sp-muted">{messages.noActivity}</p> : null}
      {activities.map((activity) => (
        <div key={activity.id} className="spm-border-b spm-border-sp-border spm-py-2 last:spm-border-b-0">
          <div className="spm-text-sm">{activity.description ?? activity.action}</div>
          <div className="spm-text-xs spm-text-sp-muted">
            {activity.timestamp ? formatDate(activity.timestamp, locale) : ""}
            {activity.actor?.displayName ? ` · ${activity.actor.displayName}` : ""}
          </div>
        </div>
      ))}
      <div className="spm-mt-3">
        <Button onClick={onClose}>{messages.cancel}</Button>
      </div>
    </Dialog>
  );
}
