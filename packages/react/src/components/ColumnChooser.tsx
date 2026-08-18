import type { ListColumn } from "@namphuongso/sharepoint-file-manager-core";
import { Button, Checkbox, Text, tokens } from "@fluentui/react-components";
import { useEffect, useRef } from "react";
import type { Messages } from "../i18n/messages";
import {
  DEFAULT_COLUMN_VISIBILITY,
  type BuiltInColumnId,
  type ColumnVisibilitySettings,
} from "../utils/column-settings";

const BUILT_IN_COLUMNS: Array<{ id: BuiltInColumnId; labelKey: keyof Messages }> = [
  { id: "modified", labelKey: "modified" },
  { id: "modifiedBy", labelKey: "modifiedBy" },
  { id: "size", labelKey: "size" },
  { id: "created", labelKey: "created" },
  { id: "createdBy", labelKey: "createdBy" },
];

export function ColumnChooser({
  open,
  messages,
  settings,
  metadataColumns,
  onChange,
  onClose,
}: {
  open: boolean;
  messages: Messages;
  settings: ColumnVisibilitySettings;
  metadataColumns: ListColumn[];
  onChange: (settings: ColumnVisibilitySettings) => void;
  onClose: () => void;
}) {
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    function handlePointer(event: MouseEvent) {
      if (ref.current && !ref.current.contains(event.target as Node)) onClose();
    }
    document.addEventListener("mousedown", handlePointer);
    return () => document.removeEventListener("mousedown", handlePointer);
  }, [open, onClose]);

  if (!open) return null;

  function toggleBuiltIn(id: BuiltInColumnId) {
    onChange({ ...settings, [id]: !settings[id] });
  }

  function toggleMetadata(name: string) {
    const current = settings.metadataColumnNames;
    const next = current.includes(name) ? current.filter((entry) => entry !== name) : [...current, name];
    onChange({ ...settings, metadataColumnNames: next });
  }

  return (
    <div
      ref={ref}
      className="spm-column-chooser"
      style={{
        background: tokens.colorNeutralBackground1,
        borderColor: tokens.colorNeutralStroke2,
        boxShadow: tokens.shadow16,
      }}
    >
      <Text size={200} weight="semibold" block style={{ marginBottom: 8, color: tokens.colorNeutralForeground3 }}>
        {messages.columns}
      </Text>
      <div className="spm-space-y-1">
        {BUILT_IN_COLUMNS.map((column) => (
          <Checkbox
            key={column.id}
            label={messages[column.labelKey]}
            checked={settings[column.id]}
            onChange={() => toggleBuiltIn(column.id)}
          />
        ))}
      </div>
      {metadataColumns.length > 0 ? (
        <>
          <div className="spm-my-2" style={{ borderTop: `1px solid ${tokens.colorNeutralStroke2}` }} />
          <Text size={200} weight="semibold" block style={{ marginBottom: 8, color: tokens.colorNeutralForeground3 }}>
            {messages.metadata}
          </Text>
          <div className="spm-max-h-40 spm-space-y-1 spm-overflow-auto">
            {metadataColumns.map((column) => (
              <Checkbox
                key={column.id}
                label={column.displayName}
                checked={settings.metadataColumnNames.includes(column.name)}
                onChange={() => toggleMetadata(column.name)}
              />
            ))}
          </div>
        </>
      ) : null}
      <div className="spm-mt-3 spm-flex spm-justify-end spm-gap-2">
        <Button
          size="small"
          onClick={() => onChange({ ...DEFAULT_COLUMN_VISIBILITY, metadataColumnNames: settings.metadataColumnNames })}
        >
          {messages.resetColumns}
        </Button>
        <Button size="small" appearance="primary" onClick={onClose}>
          {messages.save}
        </Button>
      </div>
    </div>
  );
}
