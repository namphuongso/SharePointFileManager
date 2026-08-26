const NAME_FIELD = "FileLeafRef";
const MODIFIED_FIELD = "Modified";
const SIZE_FIELD = "File_x0020_Size";

export interface ColumnLayout {
  order: string[];
  widths: Record<string, number>;
}

/** Độ rộng mặc định khớp bảng document library (px, table-layout fixed). */
export function defaultColumnWidth(internalName: string): number {
  if (internalName === NAME_FIELD) return 280;
  if (internalName === MODIFIED_FIELD) return 160;
  if (internalName === SIZE_FIELD) return 120;
  return 168;
}

export function minColumnWidth(internalName: string): number {
  if (internalName === NAME_FIELD) return 160;
  return 80;
}

/**
 * Ghép thứ tự đã lưu với cột đang hiện: giữ vị trí cũ, cột mới chen vào cuối
 * (trước Size nếu Size vẫn là cột cuối mặc định).
 */
export function mergeColumnOrder(stored: readonly string[] | undefined, visible: readonly string[]): string[] {
  const allowed = new Set(visible);
  const kept = (stored ?? []).filter((id) => allowed.has(id));
  const missing = visible.filter((id) => !kept.includes(id));
  if (missing.length === 0) return kept;

  const sizeIndex = kept.lastIndexOf(SIZE_FIELD);
  if (sizeIndex >= 0) {
    return [...kept.slice(0, sizeIndex), ...missing, ...kept.slice(sizeIndex)];
  }
  return [...kept, ...missing];
}

/** Đưa cột `fromId` trước hoặc sau `toId`. */
export function moveColumn(
  order: readonly string[],
  fromId: string,
  toId: string,
  place: "before" | "after" = "before",
): string[] {
  if (fromId === toId && place === "before") return [...order];
  const next = order.filter((id) => id !== fromId);
  const to = next.indexOf(toId);
  if (to < 0) return [...order];
  next.splice(place === "after" ? to + 1 : to, 0, fromId);
  return next;
}

/** Giữ InternalName đã ẩn trong storage để hiện lại đúng vị trí tương đối. */
export function persistColumnOrder(stored: readonly string[] | undefined, visibleOrder: readonly string[]): string[] {
  const visible = new Set(visibleOrder);
  const hidden = (stored ?? []).filter((id) => !visible.has(id));
  return [...visibleOrder, ...hidden];
}
