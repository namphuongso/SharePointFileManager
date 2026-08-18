import type { ComponentProps, ReactElement, ReactNode } from "react";
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

export function FolderIcon({ className = "spm-h-5 spm-w-5" }: { className?: string }) {
  return (
    <svg className={`${className} spm-text-[#ffb900]`} viewBox="0 0 24 24" fill="currentColor">
      <path d="M10 4l2 2h8a2 2 0 012 2v10a2 2 0 01-2 2H4a2 2 0 01-2-2V6a2 2 0 012-2h6z" />
    </svg>
  );
}

export function FileIcon() {
  return (
    <svg className="spm-h-5 spm-w-5 spm-text-slate-500" viewBox="0 0 24 24" fill="currentColor">
      <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8l-6-6zm1 7V3.5L19.5 9H15z" />
    </svg>
  );
}

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

export function CommandButton({
  children,
  onClick,
  disabled,
  primary,
  title,
  icon,
  className = "",
}: {
  children?: ReactNode;
  onClick?: () => void;
  disabled?: boolean;
  primary?: boolean;
  title?: string;
  icon?: ReactElement;
  className?: string;
}) {
  return (
    <FluentButton
      appearance={primary ? "primary" : "subtle"}
      disabled={disabled}
      onClick={onClick}
      title={title}
      icon={icon}
      className={className}
      size="medium"
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

export function ErrorBanner({ message, onRetry }: { message: string; onRetry?: () => void }) {
  return (
    <MessageBar intent="error" className="spm-mx-3 spm-my-2">
      <MessageBarBody>
        {message}
        {onRetry ? (
          <FluentButton appearance="transparent" size="small" onClick={onRetry} className="spm-ml-2">
            Retry
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
