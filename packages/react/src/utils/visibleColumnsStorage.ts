const STORAGE_PREFIX = "sp_file_manager_visible_extra";

function storageKey(scope: string): string {
  return `${STORAGE_PREFIX}:${scope}`;
}

/** Đọc InternalName cột extra đã tick. JSON hỏng / SSR → coi như chưa lưu. */
export function readVisibleExtraColumns(scope: string): string[] | undefined {
  if (typeof window === "undefined") return undefined;
  try {
    const raw = window.localStorage.getItem(storageKey(scope));
    if (!raw) return undefined;
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed) || parsed.some((name) => typeof name !== "string")) {
      return undefined;
    }
    return parsed;
  } catch {
    return undefined;
  }
}

/** Ghi set cột extra — F5 vẫn giữ ẩn/hiện; không đụng REST. */
export function writeVisibleExtraColumns(scope: string, names: readonly string[]): void {
  if (typeof window === "undefined") return;
  try {
    window.localStorage.setItem(storageKey(scope), JSON.stringify(names));
  } catch {
    /* quota / private mode — UI vẫn đổi được trong phiên hiện tại */
  }
}
