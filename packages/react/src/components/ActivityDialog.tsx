import type { DriveItemActivity } from "@namphuongso/sharepoint-file-manager-core";
import { ArrowRedoRegular, EditRegular, FolderArrowRightRegular, ShareRegular } from "@fluentui/react-icons";
import type { Messages } from "../i18n/messages";
import { Button, Dialog, formatDate, formatRelativeDate } from "./ui";

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
        <div key={activity.id} className="spm-mb-2 spm-rounded-md spm-border spm-border-sp-border spm-bg-slate-50 spm-p-3">
          <div className="spm-flex spm-items-start spm-gap-3">
            <span className="spm-mt-0.5 spm-text-sky-600">{iconForActivity(activity.action)}</span>
            <div className="spm-min-w-0">
              <div className="spm-text-sm">
                {buildActivitySentence(activity, locale, messages)}
              </div>
              <div className="spm-text-xs spm-text-sp-muted">
                {activity.timestamp ? formatRelativeDate(activity.timestamp, locale) : ""}
              </div>
              <div className="spm-text-xs spm-text-sp-muted">
                {[activity.timestamp ? formatDate(activity.timestamp, locale) : "", activity.actor?.displayName ?? ""]
                  .filter((part) => part.length > 0)
                  .join(" · ")}
              </div>
            </div>
          </div>
        </div>
      ))}
      <div className="spm-mt-3">
        <Button onClick={onClose}>{messages.cancel}</Button>
      </div>
    </Dialog>
  );
}

function buildActivitySentence(
  activity: DriveItemActivity,
  locale: string,
  messages: Messages,
): string {
  const action = (activity.action ?? "").toLowerCase();
  const itemName = extractItemName(activity);
  const target = itemName ? ` ${itemName}` : "";
  if (!locale.toLowerCase().startsWith("vi")) return (activity.description ?? action) || messages.activity;
  if (action.includes("edit")) return `Bạn đã chỉnh sửa${target}`;
  if (action.includes("create")) return `Bạn đã tạo${target}`;
  if (action.includes("move")) return `Bạn đã di chuyển${target}`;
  if (action.includes("rename")) return `Bạn đã đổi tên${target}`;
  if (action.includes("delete")) return `Bạn đã xóa${target}`;
  if (action.includes("share")) return `Bạn đã chia sẻ${target}`;
  if (action.includes("access")) return `Bạn đã truy cập${target}`;
  return activity.description ?? messages.activity;
}

function extractItemName(activity: DriveItemActivity): string {
  if (!activity.description) return "";
  const parts = activity.description.split("·").map((part) => part.trim()).filter(Boolean);
  return parts.length > 1 ? parts[1]! : "";
}

function iconForActivity(action?: string) {
  const normalized = (action ?? "").toLowerCase();
  if (normalized.includes("move")) return <FolderArrowRightRegular />;
  if (normalized.includes("share")) return <ShareRegular />;
  if (normalized.includes("edit") || normalized.includes("rename")) return <EditRegular />;
  return <ArrowRedoRegular />;
}
