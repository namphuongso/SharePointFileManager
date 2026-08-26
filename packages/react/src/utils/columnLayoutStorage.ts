import type { ColumnLayout } from "./columnLayout";

const STORAGE_PREFIX = "sp_file_manager_column_layout";

function storageKey(scope: string): string {
  return `${STORAGE_PREFIX}:${scope}`;
}

function isWidthMap(value: unknown): value is Record<string, number> {
  if (!value || typeof value !== "object") return false;
  return Object.values(value).every((width) => typeof width === "number" && Number.isFinite(width));
}

/** Đọc thứ tự + độ rộng cột. JSON hỏng / SSR → layout mặc định. */
export function readColumnLayout(scope: string): ColumnLayout | undefined {
  if (typeof window === "undefined") return undefined;
  try {
    const raw = window.localStorage.getItem(storageKey(scope));
    if (!raw) return undefined;
    const parsed: unknown = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return undefined;
    const order = (parsed as { order?: unknown }).order;
    const widths = (parsed as { widths?: unknown }).widths;
    if (!Array.isArray(order) || order.some((id) => typeof id !== "string")) return undefined;
    if (!isWidthMap(widths)) return undefined;
    return { order, widths };
  } catch {
    return undefined;
  }
}

/** Ghi layout; F5 giữ vị trí/độ rộng — không đụng REST. */
export function writeColumnLayout(scope: string, layout: ColumnLayout): void {
  if (typeof window === "undefined") return;
  try {
    window.localStorage.setItem(storageKey(scope), JSON.stringify(layout));
  } catch {
    /* quota / private mode — UI vẫn đổi được trong phiên hiện tại */
  }
}
