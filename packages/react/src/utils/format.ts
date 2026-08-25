/** Định dạng kích thước file trên cột Size. */
export function formatBytes(size?: number): string {
  if (size === undefined) return "—";
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

/** Thời gian tương đối (locale); quá 7 ngày thì ngày tháng tuyệt đối. */
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
  return new Intl.DateTimeFormat(locale, {
    weekday: "long",
    hour: "numeric",
    minute: "2-digit",
  }).format(date);
}

/** Số con folder trên cột kích thước (giống File Size SharePoint). */
export function formatItemCount(count: number | undefined, template: string): string {
  if (count === undefined) return "—";
  return template.replace("{count}", String(count));
}
