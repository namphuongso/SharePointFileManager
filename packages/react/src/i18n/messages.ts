import type { Messages } from "../types";

/** Chuỗi UI vi/en. getMessages chọn theo locale, có thể ghi đè từng key. */

const vi: Messages = {
  files: "Tệp",
  name: "Tên",
  modified: "Sửa đổi",
  size: "Kích thước",
  empty: "Thư mục trống",
  emptyHint: "Không có tệp hoặc thư mục nào trong thư mục này.",
  retry: "Thử lại",
  refresh: "Làm mới",
  unknownError: "Đã xảy ra lỗi.",
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
  unknownError: "Something went wrong.",
};

export function getMessages(locale: string, overrides?: Partial<Messages>): Messages {
  const base = locale.toLowerCase().startsWith("en") ? en : vi;
  return overrides ? { ...base, ...overrides } : base;
}
