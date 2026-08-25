import type { ListChildrenSort } from "@namphuongso/sharepoint-file-manager-core";

const STORAGE_PREFIX = "sp_file_manager_column_sort";

function storageKey(scope: string): string {
  return `${STORAGE_PREFIX}:${scope}`;
}

function isSortDirection(value: unknown): value is ListChildrenSort["direction"] {
  return value === "asc" || value === "desc";
}

/** Đọc sort cột đã chọn. JSON hỏng / SSR → không sort (mặc định SharePoint). */
export function readColumnSort(scope: string): ListChildrenSort | undefined {
  if (typeof window === "undefined") return undefined;
  try {
    const raw = window.localStorage.getItem(storageKey(scope));
    if (!raw) return undefined;
    const parsed: unknown = JSON.parse(raw);
    if (
      !parsed ||
      typeof parsed !== "object" ||
      typeof (parsed as { field?: unknown }).field !== "string" ||
      !isSortDirection((parsed as { direction?: unknown }).direction)
    ) {
      return undefined;
    }
    return {
      field: (parsed as ListChildrenSort).field,
      direction: (parsed as ListChildrenSort).direction,
    };
  } catch {
    return undefined;
  }
}

/** Ghi sort; `undefined` = bỏ lọc, xóa key để F5 không còn $orderby tùy chỉnh. */
export function writeColumnSort(scope: string, sort: ListChildrenSort | undefined): void {
  if (typeof window === "undefined") return;
  try {
    const key = storageKey(scope);
    if (!sort) {
      window.localStorage.removeItem(key);
      return;
    }
    window.localStorage.setItem(key, JSON.stringify(sort));
  } catch {
    /* quota / private mode — UI vẫn sort được trong phiên hiện tại */
  }
}
