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
  /** Không có OpenItems trên item đang thao tác (mở file / tải file·folder). */
  noOpenPermission: string;
  noOpenPermissionHint: string;
  /** Không có AddListItems trên folder hiện tại. */
  noAddPermission: string;
  /** Lỗi GET URL mở file (không phải 403). */
  openFileError: string;
  /** aria-label / title dòng file — gợi ý bấm để mở. */
  openFile: string;
  /** Menu thao tác folder — mở / vào thư mục. */
  openFolder: string;
  /** Nút ⋯ trên dòng (hover). */
  moreActions: string;
  /** Menu chuột phải / action tải file xuống. */
  download: string;
  /** Đang GET /$value hoặc chuẩn bị zip. */
  downloading: string;
  /** Đang liệt kê cây folder trước khi tải. */
  downloadingFolderListing: string;
  /** Đang tải từng file vào zip — `{done}` `{total}`. */
  downloadingFolderFiles: string;
  /** Đang nén zip phía client. */
  downloadingFolderZipping: string;
  /** Toast — tải xong. */
  downloadSuccess: string;
  /** Lỗi tải nhị phân (không phải 403). */
  downloadError: string;
  /** Nút / dialog tạo thư mục. */
  newFolder: string;
  folderName: string;
  createFolder: string;
  createFolderError: string;
  /** Menu New (toolbar + chuột phải). */
  newItem: string;
  uploadFiles: string;
  uploadFolder: string;
  newWord: string;
  newExcel: string;
  newPowerPoint: string;
  fileName: string;
  createDocumentError: string;
  /** Nút upload file. */
  upload: string;
  uploadError: string;
  /** Đang POST upload (banner trạng thái). */
  uploading: string;
  /** Tên trùng khi tạo / upload (Conflict). */
  nameConflict: string;
  /** File quá lớn (413). */
  fileTooLarge: string;
  /** Tên không hợp lệ (ký tự cấm). */
  invalidName: string;
  /** Toast — upload 1 file xong. */
  uploadSuccess: string;
  /** Toast — upload cả thư mục xong. */
  uploadFolderSuccess: string;
  /** Toast — tạo thư mục xong. */
  createFolderSuccess: string;
  /** Toast — tạo file trống (Word/Excel/PowerPoint) xong. */
  createDocumentSuccess: string;
  cancel: string;
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
