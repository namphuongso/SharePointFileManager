import { useId, type FormEvent, type ReactElement } from "react";
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Spinner,
} from "@fluentui/react-components";
import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../../types/messages";

export interface DeleteItemDialogProps {
  open: boolean;
  item: SharePointItem | null;
  onOpenChange: (open: boolean) => void;
  onConfirm: () => Promise<void>;
  messages: Messages;
  isPending?: boolean;
}

/** Xác nhận soft-delete — tên item + Hủy / Xóa. */
export function DeleteItemDialog({
  open,
  item,
  onOpenChange,
  onConfirm,
  messages,
  isPending = false,
}: DeleteItemDialogProps): ReactElement {
  const titleId = useId();

  async function handleSubmit(event: FormEvent) {
    event.preventDefault();
    if (!item || isPending) return;
    await onConfirm();
  }

  const confirmLabel = item?.type === "folder" ? messages.deleteFolder : messages.deleteFile;
  const body = messages.deleteConfirmBody.replace("{name}", item?.name ?? "");

  return (
    <Dialog open={open} onOpenChange={(_, data) => onOpenChange(data.open)}>
      <DialogSurface aria-labelledby={titleId}>
        <form onSubmit={(e) => void handleSubmit(e)}>
          <DialogBody>
            <DialogTitle id={titleId}>{messages.deleteConfirmTitle}</DialogTitle>
            <DialogContent>{body}</DialogContent>
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
                disabled={isPending || !item}
                icon={isPending ? <Spinner size="tiny" /> : undefined}
              >
                {confirmLabel}
              </Button>
            </DialogActions>
          </DialogBody>
        </form>
      </DialogSurface>
    </Dialog>
  );
}
