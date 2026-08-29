export interface Messages {
  files: string;
  name: string;
  modified: string;
  size: string;
  empty: string;
  emptyHint: string;
  /** Không có ViewListItems trên folder hiện tại. */
  noViewPermission: string;
  noViewPermissionHint: string;
  retry: string;
  refresh: string;
  loadMore: string;
  unknownError: string;
  /** Nhãn nút mở column picker — ẩn/hiện cột. */
  columns: string;
  /** Cột Name/Modified luôn hiện trên bảng. */
  alwaysVisible: string;
  /** Cột tùy chọn trong picker. */
  moreColumns: string;
  /** Folder trên cột kích thước: "2 khoản mục". */
  itemCount: string;
  /** Nhãn nút đổi ngôn ngữ hiển thị. */
  language: string;
  /** aria-sort: tăng dần. */
  sortAscending: string;
  /** aria-sort: giảm dần. */
  sortDescending: string;
  /** Sort chữ / người: A → Z. */
  sortAZ: string;
  /** Sort chữ / người: Z → A. */
  sortZA: string;
  /** Sort ngày: cũ trước. */
  sortOldest: string;
  /** Sort ngày: mới trước. */
  sortNewest: string;
  /** Sort số: nhỏ trước. */
  sortSmallest: string;
  /** Sort số: lớn trước. */
  sortLargest: string;
  /** aria-label tay cầm kéo rộng cột. */
  resizeColumn: string;
  /** Tab browse thư viện theo folder. */
  tabLibrary: string;
  /** Tab Search — item user được xem / được chia sẻ trong thư viện. */
  tabAccessible: string;
  /** Empty state tab accessible. */
  searchEmpty: string;
  searchEmptyHint: string;
  /**
   * Nhãn cột ghi đè theo InternalName.
   * Mặc định UI dùng Title SharePoint trả theo locale — chỉ đặt khi muốn ghi đè.
   */
  fieldLabels: Record<string, string>;
}
