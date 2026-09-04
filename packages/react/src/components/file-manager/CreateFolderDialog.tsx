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
import type { Messages } from "../../types/messages";

export interface CreateFolderDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onSubmit: (name: string) => Promise<void>;
  messages: Messages;
  isPending?: boolean;
  errorMessage?: string;
}

/** Dialog nhập tên thư mục mới — Fluent Dialog, không card tùy biến. */
export function CreateFolderDialog({
  open,
  onOpenChange,
  onSubmit,
  messages,
  isPending = false,
  errorMessage,
}: CreateFolderDialogProps) {
  const inputId = useId();
  const [name, setName] = useState("");

  useEffect(() => {
    if (open) setName("");
  }, [open]);

  async function handleSubmit(event: FormEvent) {
    event.preventDefault();
    const trimmed = name.trim();
    if (!trimmed || isPending) return;
    await onSubmit(trimmed);
  }

  return (
    <Dialog open={open} onOpenChange={(_, data) => onOpenChange(data.open)}>
      <DialogSurface>
        <form onSubmit={(e) => void handleSubmit(e)}>
          <DialogBody>
            <DialogTitle>{messages.newFolder}</DialogTitle>
            <DialogContent>
              <Field
                label={messages.folderName}
                required
                validationState={errorMessage ? "error" : "none"}
                validationMessage={errorMessage}
              >
                <Input
                  id={inputId}
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
