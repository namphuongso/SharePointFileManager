import { useEffect, useState } from "react";
import type { ListColumn, SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../../i18n/messages";
import { Button, Dialog } from "../ui";

export function BulkMetadataDialog({
  items,
  columns,
  open,
  messages,
  pending,
  onClose,
  onSubmit,
}: {
  items: SharePointItem[];
  columns: ListColumn[];
  open: boolean;
  messages: Messages;
  pending?: boolean;
  onClose: () => void;
  onSubmit: (fields: Record<string, string | number | boolean | null>) => void;
}) {
  const editable = columns.filter((column) => !column.readOnly);
  const [values, setValues] = useState<Record<string, string>>({});

  useEffect(() => {
    if (!open) {
      setValues({});
      return;
    }
    const initial: Record<string, string> = {};
    for (const column of editable) initial[column.name] = "";
    setValues(initial);
  }, [open, columns]);

  if (editable.length === 0) return null;

  return (
    <Dialog
      open={open}
      title={`${messages.bulkEditMetadata} (${items.length})`}
      onClose={onClose}
      footer={
        <>
          <Button onClick={onClose}>{messages.cancel}</Button>
          <Button
            variant="primary"
            disabled={pending}
            onClick={() => {
              const payload: Record<string, string | number | boolean | null> = {};
              for (const column of editable) {
                const raw = values[column.name]?.trim() ?? "";
                if (raw) payload[column.name] = raw;
              }
              if (Object.keys(payload).length === 0) return;
              onSubmit(payload);
            }}
          >
            {messages.save}
          </Button>
        </>
      }
    >
      <p className="spm-mb-3 spm-text-sm spm-text-sp-muted">{messages.bulkEditMetadataHint}</p>
      {editable.map((column) => (
        <label key={column.id} className="spm-mb-3 spm-block spm-text-sm">
          <span className="spm-mb-1 spm-block spm-text-sp-muted">{column.displayName}</span>
          <input
            className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-2 spm-py-1.5"
            value={values[column.name] ?? ""}
            placeholder={messages.leaveBlankToSkip}
            onChange={(event) => setValues((current) => ({ ...current, [column.name]: event.target.value }))}
          />
        </label>
      ))}
    </Dialog>
  );
}
