import { useEffect, useId, useState, type FormEvent } from "react";
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Field,
  Input,
  Spinner,
} from "@fluentui/react-components";
import type { NewDocumentKind } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../../types/messages";

export interface CreateDocumentDialogProps {
  open: boolean;
  kind: NewDocumentKind | null;
  onOpenChange: (open: boolean) => void;
  onSubmit: (values: { name: string }) => Promise<void>;
  messages: Messages;
  defaultName: string;
  isPending?: boolean;
  errorMessage?: string;
}

function dialogTitle(kind: NewDocumentKind, messages: Messages): string {
  const titleByKind = {
    word: messages.newWord,
    excel: messages.newExcel,
    powerpoint: messages.newPowerPoint,
  } as const satisfies Record<NewDocumentKind, string>;
  return titleByKind[kind];
}

/** Dialog tên file khi tạo tài liệu Office trống. */
export function CreateDocumentDialog({
  open,
  kind,
  onOpenChange,
  onSubmit,
  messages,
  defaultName,
  isPending = false,
  errorMessage,
}: CreateDocumentDialogProps) {
  const nameId = useId();
  const [name, setName] = useState(defaultName);

  useEffect(() => {
    if (open) setName(defaultName);
  }, [open, defaultName, kind]);

  if (!kind) return null;

  async function handleSubmit(event: FormEvent) {
    event.preventDefault();
    const trimmed = name.trim();
    if (!trimmed || isPending) return;
    await onSubmit({ name: trimmed });
  }

  return (
    <Dialog open={open} onOpenChange={(_, data) => onOpenChange(data.open)}>
      <DialogSurface>
        <form onSubmit={(e) => void handleSubmit(e)}>
          <DialogBody>
            <DialogTitle>{dialogTitle(kind, messages)}</DialogTitle>
            <DialogContent>
              <Field
                label={messages.fileName}
                required
                validationState={errorMessage ? "error" : "none"}
                validationMessage={errorMessage}
              >
                <Input
                  id={nameId}
                  value={name}
                  onChange={(_, data) => setName(data.value)}
                  disabled={isPending}
                  autoFocus
                />
              </Field>
            </DialogContent>
            <DialogActions>
              <Button
                appearance="secondary"
                type="button"
                disabled={isPending}
                onClick={() => onOpenChange(false)}
              >
                {messages.cancel}
              </Button>
              <Button
                appearance="primary"
                type="submit"
                disabled={isPending || !name.trim()}
                icon={isPending ? <Spinner size="tiny" /> : undefined}
              >
                {messages.createFolder}
              </Button>
            </DialogActions>
          </DialogBody>
        </form>
      </DialogSurface>
    </Dialog>
  );
}
