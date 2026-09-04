/** Tab browse vs Search REST — đồng bộ query `view`. */
export type BrowserViewParam = "library" | "accessible";

export interface FileBrowserUrlState {
  view: BrowserViewParam;
  /** UniqueId folder hiện tại; thiếu / "root" = gốc thư viện. */
  folderId: string | undefined;
}

const VIEW_PARAM = "view";
const FOLDER_PARAM = "folder";

/** Đọc ?view=&folder= từ thanh địa chỉ. */
export function readFileBrowserUrl(search = window.location.search): FileBrowserUrlState {
  const params = new URLSearchParams(search);
  const view: BrowserViewParam =
    params.get(VIEW_PARAM) === "accessible" ? "accessible" : "library";
  const raw = params.get(FOLDER_PARAM)?.trim();
  const folderId = raw && raw.toLowerCase() !== "root" ? raw.replace(/[{}]/g, "") : undefined;
  return { view, folderId };
}

/**
 * Ghi ?view=&folder= (giữ param khác của host).
 * push = điều hướng / back; replace = hydrate hoặc sửa URL lỗi.
 */
export function writeFileBrowserUrl(
  state: FileBrowserUrlState,
  mode: "push" | "replace",
  rootFolderIds: ReadonlySet<string> = new Set(["root"]),
): void {
  const url = new URL(window.location.href);

  if (state.view === "accessible") {
    url.searchParams.set(VIEW_PARAM, "accessible");
  } else {
    url.searchParams.delete(VIEW_PARAM);
  }

  const folder = state.folderId?.replace(/[{}]/g, "");
  const isRoot = !folder || rootFolderIds.has(folder.toLowerCase()) || rootFolderIds.has(folder);
  if (isRoot) {
    url.searchParams.delete(FOLDER_PARAM);
  } else {
    url.searchParams.set(FOLDER_PARAM, folder);
  }

  const next = `${url.pathname}${url.search}${url.hash}`;
  const current = `${window.location.pathname}${window.location.search}${window.location.hash}`;
  if (next === current) return;

  if (mode === "push") {
    window.history.pushState(null, "", next);
  } else {
    window.history.replaceState(null, "", next);
  }
}
