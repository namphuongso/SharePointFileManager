import type {
  DriveItemActivity,
  SharePointItem,
  SharePointPermission,
} from "@namphuongso/sharepoint-file-manager-core";
import { Button, Divider, Text, tokens } from "@fluentui/react-components";
import {
  DismissRegular,
  HistoryRegular,
  LockClosedRegular,
  PeopleRegular,
} from "@fluentui/react-icons";
import type { Messages } from "../i18n/messages";
import { FileTypeIcon } from "./FileTypeIcon";
import { formatBytes, formatDate } from "./ui";
import { useEffect, useRef } from "react";

export function DetailsPane({
  item,
  locale,
  messages,
  onClose,
  onOpenProperties,
  onPreview,
  onShare,
  permissions,
  permissionsLoading,
  activities,
  activitiesLoading,
  onOpenManageAccess,
  onOpenActivity,
}: {
  item: SharePointItem;
  locale: string;
  messages: Messages;
  onClose: () => void;
  onOpenProperties: () => void;
  onPreview?: () => void;
  onShare?: () => void;
  permissions?: SharePointPermission[];
  permissionsLoading?: boolean;
  activities?: DriveItemActivity[];
  activitiesLoading?: boolean;
  onOpenManageAccess?: () => void;
  onOpenActivity?: () => void;
}) {
  const closeButtonRef = useRef<HTMLButtonElement>(null);
  const titleId = `spm-details-title-${item.id}`;
  useEffect(() => {
    closeButtonRef.current?.focus();
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") onClose();
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, [onClose]);
  return (
    <aside
      className="spm-details-pane spm-flex spm-flex-col"
      role="complementary"
      aria-labelledby={titleId}
      style={{ background: tokens.colorNeutralBackground1 }}
    >
      <div
        className="spm-details-header"
        style={{ background: tokens.colorNeutralBackground2 }}
      >
        <Text id={titleId} weight="semibold">{messages.details}</Text>
        <Button ref={closeButtonRef} appearance="subtle" icon={<DismissRegular />} onClick={onClose} aria-label={messages.cancel} />
      </div>
      <div className="spm-flex-1 spm-space-y-4 spm-p-4">
        {item.thumbnailUrl && item.type === "file" ? (
          <button type="button" className="spm-details-preview" onClick={onPreview} disabled={!onPreview}>
            <img src={item.thumbnailUrl} alt={item.name} />
          </button>
        ) : null}
        <div className="spm-flex spm-items-start spm-gap-3">
          <FileTypeIcon item={item} size="lg" />
          <div className="spm-min-w-0 spm-flex-1">
            <Text weight="semibold" block wrap>
              {item.name}
            </Text>
            <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
              {item.type === "folder" ? messages.folder : messages.file}
            </Text>
          </div>
        </div>
        {item.capabilities?.isCheckedOut ? (
          <div
            className="spm-flex spm-items-center spm-gap-2 spm-rounded-sm spm-px-3 spm-py-2"
            style={{ background: tokens.colorPaletteYellowBackground1, color: tokens.colorPaletteDarkOrangeForeground1 }}
          >
            <LockClosedRegular />
            <Text size={200}>
              {messages.checkedOut}
              {item.capabilities.checkedOutBy?.displayName
                ? `: ${item.capabilities.checkedOutBy.displayName}`
                : ""}
            </Text>
          </div>
        ) : null}
        {item.sensitivityLabel ? (
          <Text size={200}>
            <Text style={{ color: tokens.colorNeutralForeground3 }}>{messages.sensitivityLabel}: </Text>
            {item.sensitivityLabel}
          </Text>
        ) : null}
        <Divider />
        <dl className="spm-space-y-3">
          <DetailRow label={messages.modified} value={formatDate(item.lastModifiedDateTime, locale)} />
          <DetailRow
            label={messages.modifiedBy}
            value={item.lastModifiedBy?.displayName ?? item.lastModifiedBy?.email ?? "—"}
          />
          <DetailRow label={messages.created} value={formatDate(item.createdDateTime, locale)} />
          <DetailRow
            label={messages.createdBy}
            value={item.createdBy?.displayName ?? item.createdBy?.email ?? "—"}
          />
          {item.type === "file" ? <DetailRow label={messages.size} value={formatBytes(item.size)} /> : null}
          {item.mimeType ? <DetailRow label={messages.type} value={item.mimeType} /> : null}
        </dl>
        {item.metadata && Object.keys(item.metadata).length > 0 ? (
          <>
            <Divider />
            <div className="spm-space-y-2">
              <Text size={200} weight="semibold" style={{ color: tokens.colorNeutralForeground3, textTransform: "uppercase" }}>
                {messages.metadata}
              </Text>
              {Object.entries(item.metadata)
                .filter(([key]) => !isNoisyMetadataKey(key))
                .slice(0, 6)
                .map(([key, value]) => (
                  <DetailRow
                    key={key}
                    label={humanizeMetadataKey(key)}
                    value={value === null || value === undefined || value === "" ? "—" : String(value)}
                  />
                ))}
            </div>
          </>
        ) : null}
        <Divider />
        <div className="spm-space-y-2">
          <Text size={300} weight="semibold" style={{ color: tokens.colorNeutralForeground2 }}>
            {messages.activityLog}
          </Text>
          {activitiesLoading ? (
            <Text size={200}>{messages.loading}</Text>
          ) : activities && activities.length > 0 ? (
            <div className="spm-space-y-2">
              {activities.slice(0, 3).map((activity) => (
                <div key={activity.id} className="spm-rounded-sm spm-p-2" style={{ background: tokens.colorNeutralBackground2 }}>
                  <Text size={200} block>{formatActivitySummary(activity, locale, messages)}</Text>
                  <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>
                    {[activity.timestamp ? formatDate(activity.timestamp, locale) : "", activity.actor?.displayName ?? ""]
                      .filter((part) => part.length > 0)
                      .join(" · ")}
                  </Text>
                </div>
              ))}
            </div>
          ) : (
            <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>{messages.noActivity}</Text>
          )}
          {onOpenActivity ? (
            <Button appearance="subtle" icon={<HistoryRegular />} onClick={onOpenActivity}>
              {messages.activity}
            </Button>
          ) : null}
        </div>
        <Divider />
        <div className="spm-space-y-2">
          <Text size={300} weight="semibold" style={{ color: tokens.colorNeutralForeground2 }}>
            {messages.manageAccess}
          </Text>
          {permissionsLoading ? (
            <Text size={200}>{messages.loading}</Text>
          ) : (
            <div className="spm-details-access-summary">
              <Text size={200}>
                {(permissions ?? []).filter((permission) => permission.kind === "user" || permission.kind === "group" || permission.kind === "siteGroup").length}
              </Text>
              <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                {messages.people}
              </Text>
            </div>
          )}
          {onOpenManageAccess ? (
            <Button appearance="subtle" icon={<PeopleRegular />} onClick={onOpenManageAccess}>
              {messages.manageAccess}
            </Button>
          ) : null}
        </div>
        <Divider />
        <div className="spm-flex spm-flex-wrap spm-gap-2">
          {item.type === "file" && onPreview ? (
            <Button appearance="primary" onClick={onPreview}>
              {messages.preview}
            </Button>
          ) : null}
          {onShare ? <Button appearance="subtle" onClick={onShare}>{messages.share}</Button> : null}
          <Button appearance="outline" onClick={onOpenProperties}>
            {messages.properties}
          </Button>
        </div>
      </div>
    </aside>
  );
}

function humanizeMetadataKey(value: string): string {
  return value
    .replace(/_/g, " ")
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .trim();
}

function isNoisyMetadataKey(key: string): boolean {
  const normalized = key.trim();
  const lower = normalized.toLowerCase();
  return (
    normalized.startsWith("@") ||
    normalized.startsWith("_") ||
    normalized.startsWith("OData__") ||
    lower === "id" ||
    lower.endsWith("lookupid") ||
    lower.endsWith("stringid") ||
    lower.includes("etag")
  );
}

function formatActivitySummary(
  activity: DriveItemActivity,
  locale: string,
  messages: Messages,
): string {
  const action = (activity.action ?? "").toLowerCase();
  const vi = locale.toLowerCase().startsWith("vi");
  if (vi) {
    if (action.includes("edit")) return "Đã chỉnh sửa";
    if (action.includes("create")) return "Đã tạo";
    if (action.includes("move")) return "Đã di chuyển";
    if (action.includes("rename")) return "Đã đổi tên";
    if (action.includes("delete")) return "Đã xóa";
    if (action.includes("share")) return "Đã chia sẻ";
    if (action.includes("access")) return "Đã truy cập";
    return activity.description ?? messages.activity;
  }
  return (activity.description ?? action) || messages.activity;
}

function DetailRow({ label, value }: { label: string; value: string }) {
  return (
    <div>
      <Text size={200} weight="semibold" block style={{ color: tokens.colorNeutralForeground3, textTransform: "uppercase" }}>
        {label}
      </Text>
      <Text block wrap>
        {value}
      </Text>
    </div>
  );
}
