export interface Messages {
  files: string;
  name: string;
  modified: string;
  size: string;
  empty: string;
  emptyHint: string;
  retry: string;
  refresh: string;
  loadMore: string;
  unknownError: string;
  /** Nhãn nút mở column picker — ẩn/hiện cột. */
  columns: string;
  /** Folder trên cột kích thước: "2 khoản mục". */
  itemCount: string;
  /** Nhãn nút đổi ngôn ngữ hiển thị. */
  language: string;
  /**
 * Nhãn cột ghi đè theo InternalName.
   * Mặc định UI dùng Title SharePoint trả theo locale — chỉ đặt khi muốn ghi đè.
   */
  fieldLabels: Record<string, string>;
}
