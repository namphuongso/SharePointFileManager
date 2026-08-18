import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import { Button, Divider, Text, tokens } from "@fluentui/react-components";
import { DismissRegular, LockClosedRegular } from "@fluentui/react-icons";
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
}: {
  item: SharePointItem;
  locale: string;
  messages: Messages;
  onClose: () => void;
  onOpenProperties: () => void;
  onPreview?: () => void;
  onShare?: () => void;
}) {
  const closeButtonRef = useRef<HTMLButtonElement>(null);
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
      role="dialog"
      aria-modal="true"
      aria-label={messages.details}
      style={{ background: tokens.colorNeutralBackground1 }}
    >
      <div
        className="spm-details-header"
        style={{ background: tokens.colorNeutralBackground2 }}
      >
        <Text weight="semibold">{messages.details}</Text>
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
        <Divider />
        <div className="spm-flex spm-flex-wrap spm-gap-2">
          {item.type === "file" && onPreview ? (
            <Button appearance="primary" onClick={onPreview}>
              {messages.preview}
            </Button>
          ) : null}
          {onShare ? <Button onClick={onShare}>{messages.share}</Button> : null}
          <Button appearance="secondary" onClick={onOpenProperties}>
            {messages.properties}
          </Button>
        </div>
      </div>
    </aside>
  );
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
