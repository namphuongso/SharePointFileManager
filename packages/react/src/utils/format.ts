const SIZE_UNITS = ["B", "KB", "MB", "GB", "TB"];

/** Định dạng kích thước file trên cột Size theo locale (giống SharePoint). */
export function formatBytes(size?: number, locale = "en"): string {
  if (size === undefined) return "—";
  if (size < 1024) return `${size} B`;
  const decimals = size < 1024 * 1024 ? 1 : 1;
  const formatter = new Intl.NumberFormat(locale, {
    minimumFractionDigits: 0,
    maximumFractionDigits: decimals,
  });
  const index = Math.min(
    Math.floor(Math.log(size) / Math.log(1024)),
    SIZE_UNITS.length - 1,
  );
  const value = size / 1024 ** index;
  return `${formatter.format(value)} ${SIZE_UNITS[index]}`;
}

/** Ngày giờ tuyệt đối theo locale (giống SharePoint). */
export function formatDate(value: string | undefined, locale: string): string {
  if (!value) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "—";
  return new Intl.DateTimeFormat(locale, {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(date);
}

/** Số con folder trên cột kích thước (giống File Size SharePoint). */
export function formatItemCount(count: number | undefined, template: string): string {
  if (count === undefined) return "—";
  return template.replace("{count}", String(count));
}
