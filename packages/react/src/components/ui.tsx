import type { ComponentProps, ReactNode } from "react";
import {
  Button as FluentButton,
  Dialog as FluentDialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  tokens,
} from "@fluentui/react-components";
import { DismissRegular } from "@fluentui/react-icons";

export function Dialog({
  title,
  open,
  onClose,
  children,
  footer,
}: {
  title: string;
  open: boolean;
  onClose: () => void;
  children: ReactNode;
  footer?: ReactNode;
}) {
  return (
    <FluentDialog open={open} onOpenChange={(_, data) => !data.open && onClose()} modalType="modal">
      <DialogSurface>
        <DialogBody>
          <DialogTitle action={<FluentButton appearance="subtle" icon={<DismissRegular />} onClick={onClose} />}>
            {title}
          </DialogTitle>
          <DialogContent>{children}</DialogContent>
          {footer ? <DialogActions>{footer}</DialogActions> : null}
        </DialogBody>
      </DialogSurface>
    </FluentDialog>
  );
}

export function Button({
  children,
  onClick,
  variant = "secondary",
  disabled,
  type = "button",
  className = "",
}: {
  children: ReactNode;
  onClick?: () => void;
  variant?: "primary" | "secondary" | "danger";
  disabled?: boolean;
  type?: "button" | "submit";
  className?: string;
}) {
  const appearance =
    variant === "primary" ? "primary" : variant === "danger" ? "primary" : "secondary";
  return (
    <FluentButton
      type={type}
      appearance={appearance}
      disabled={disabled}
      onClick={onClick}
      className={className}
      style={variant === "danger" ? { backgroundColor: tokens.colorPaletteRedBackground3 } : undefined}
    >
      {children}
    </FluentButton>
  );
}

export function TextField({
  label,
  value,
  onChange,
  placeholder,
  type = "text",
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  type?: ComponentProps<typeof Input>["type"];
}) {
  return (
    <Field label={label}>
      <Input
        type={type}
        value={value}
        placeholder={placeholder}
        onChange={(_, data) => onChange(data.value)}
      />
    </Field>
  );
}

export function ErrorBanner({
  message,
  onRetry,
  retryLabel = "Retry",
}: {
  message: string;
  onRetry?: () => void;
  retryLabel?: string;
}) {
  return (
    <MessageBar intent="error" className="spm-mx-3 spm-my-2">
      <MessageBarBody>
        {message}
        {onRetry ? (
          <FluentButton appearance="transparent" size="small" onClick={onRetry} className="spm-ml-2">
            {retryLabel}
          </FluentButton>
        ) : null}
      </MessageBarBody>
    </MessageBar>
  );
}

export function formatBytes(size?: number): string {
  if (size === undefined) return "—";
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

export function formatDate(value: string | undefined, locale: string): string {
  if (!value) return "—";
  return new Intl.DateTimeFormat(locale, { dateStyle: "short", timeStyle: "short" }).format(new Date(value));
}

export function formatRelativeDate(value: string | undefined, locale: string): string {
  if (!value) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "—";

  const diffMs = date.getTime() - Date.now();
  const absMs = Math.abs(diffMs);
  const minute = 60_000;
  const hour = 60 * minute;
  const day = 24 * hour;
  const rtf = new Intl.RelativeTimeFormat(locale, { numeric: "auto" });

  if (absMs < minute) return rtf.format(Math.round(diffMs / 1000), "second");
  if (absMs < hour) return rtf.format(Math.round(diffMs / minute), "minute");
  if (absMs < day) return rtf.format(Math.round(diffMs / hour), "hour");
  if (absMs < 7 * day) return rtf.format(Math.round(diffMs / day), "day");
  return new Intl.DateTimeFormat(locale, { day: "numeric", month: "short", year: "numeric" }).format(date);
}
