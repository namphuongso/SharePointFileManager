import type { ReactNode } from "react";

export function FolderIcon() {
  return (
    <svg className="spm-h-5 spm-w-5 spm-text-amber-500" viewBox="0 0 24 24" fill="currentColor">
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
  if (!open) return null;
  return (
    <div className="spm-fixed spm-inset-0 spm-z-50 spm-flex spm-items-center spm-justify-center spm-bg-black/40 spm-p-4">
      <div className="spm-w-full spm-max-w-lg spm-rounded-xl spm-bg-white spm-shadow-xl">
        <div className="spm-flex spm-items-center spm-justify-between spm-border-b spm-border-sp-border spm-px-4 spm-py-3">
          <h2 className="spm-text-base spm-font-semibold">{title}</h2>
          <button type="button" className="spm-text-sp-muted" onClick={onClose}>
            ×
          </button>
        </div>
        <div className="spm-max-h-[70vh] spm-overflow-auto spm-px-4 spm-py-3">{children}</div>
        {footer ? <div className="spm-flex spm-justify-end spm-gap-2 spm-border-t spm-border-sp-border spm-px-4 spm-py-3">{footer}</div> : null}
      </div>
    </div>
  );
}

export function Button({
  children,
  onClick,
  variant = "secondary",
  disabled,
  type = "button",
}: {
  children: ReactNode;
  onClick?: () => void;
  variant?: "primary" | "secondary" | "danger";
  disabled?: boolean;
  type?: "button" | "submit";
}) {
  const styles =
    variant === "primary"
      ? "spm-bg-sp-primary spm-text-white"
      : variant === "danger"
        ? "spm-bg-sp-danger spm-text-white"
        : "spm-border spm-border-sp-border spm-bg-white";
  return (
    <button
      type={type}
      disabled={disabled}
      onClick={onClick}
      className={`spm-rounded-md spm-px-3 spm-py-1.5 spm-text-sm disabled:spm-opacity-50 ${styles}`}
    >
      {children}
    </button>
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
  type?: string;
}) {
  return (
    <label className="spm-block spm-text-sm">
      <span className="spm-mb-1 spm-block spm-text-sp-muted">{label}</span>
      <input
        type={type}
        value={value}
        placeholder={placeholder}
        onChange={(event) => onChange(event.target.value)}
        className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-3 spm-py-2"
      />
    </label>
  );
}

export function ErrorBanner({ message, onRetry }: { message: string; onRetry?: () => void }) {
  return (
    <div className="spm-m-3 spm-rounded-md spm-border spm-border-red-200 spm-bg-red-50 spm-px-3 spm-py-2 spm-text-sm spm-text-red-700">
      {message}
      {onRetry ? (
        <button type="button" className="spm-ml-2 spm-underline" onClick={onRetry}>
          Retry
        </button>
      ) : null}
    </div>
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
