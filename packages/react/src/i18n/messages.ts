import type { Messages } from "../types";

/**
 * Chuỗi UI vi/en. getMessages chọn theo locale, có thể ghi đè từng key.
 * Nhãn cột không hardcode ở đây: mặc định UI dùng Title SharePoint trả theo locale
 * (config.locale, header Accept-Language); fieldLabels dành cho host muốn ghi đè từng cột.
 */

const vi: Messages = {
  files: "Tệp",
  name: "Tên",
  modified: "Sửa đổi",
  size: "Kích thước",
  empty: "Thư mục trống",
  emptyHint: "Không có tệp hoặc thư mục nào trong thư mục này.",
  retry: "Thử lại",
  refresh: "Làm mới",
  loadMore: "Tải thêm",
  unknownError: "Đã xảy ra lỗi.",
  columns: "Cột",
  alwaysVisible: "Luôn hiện",
  moreColumns: "Thêm cột",
  itemCount: "{count} khoản mục",
  language: "Ngôn ngữ",
  sortAscending: "Tăng dần",
  sortDescending: "Giảm dần",
  sortAZ: "A đến Z",
  sortZA: "Z đến A",
  sortOldest: "Cũ nhất trước",
  sortNewest: "Mới nhất trước",
  sortSmallest: "Nhỏ đến lớn",
  sortLargest: "Lớn đến nhỏ",
  resizeColumn: "Điều chỉnh độ rộng cột",
  fieldLabels: {},
};

const en: Messages = {
  files: "Files",
  name: "Name",
  modified: "Modified",
  size: "Size",
  empty: "This folder is empty",
  emptyHint: "There are no files or folders in this folder.",
  retry: "Retry",
  refresh: "Refresh",
  loadMore: "Load more",
  unknownError: "Something went wrong.",
  columns: "Columns",
  alwaysVisible: "Always shown",
  moreColumns: "More columns",
  itemCount: "{count} items",
  language: "Language",
  sortAscending: "Ascending",
  sortDescending: "Descending",
  sortAZ: "A to Z",
  sortZA: "Z to A",
  sortOldest: "Older to newer",
  sortNewest: "Newer to older",
  sortSmallest: "Smaller to larger",
  sortLargest: "Larger to smaller",
  resizeColumn: "Resize column",
  fieldLabels: {},
};

export function getMessages(locale: string, overrides?: Partial<Messages>): Messages {
  const base = locale.toLowerCase().startsWith("en") ? en : vi;
  if (!overrides) return base;
  return {
    ...base,
    ...overrides,
    fieldLabels: { ...base.fieldLabels, ...overrides.fieldLabels },
  };
}

export function fieldLabel(messages: Messages, internalName: string, fallback: string): string {
  return messages.fieldLabels[internalName] ?? fallback;
}
