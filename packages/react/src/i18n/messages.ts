export interface Messages {
  files: string;
  search: string;
  searchPlaceholder: string;
  newFolder: string;
  upload: string;
  list: string;
  grid: string;
  name: string;
  modified: string;
  size: string;
  empty: string;
  loading: string;
  retry: string;
  cancel: string;
  save: string;
  create: string;
  delete: string;
  rename: string;
  copy: string;
  move: string;
  download: string;
  share: string;
  manageAccess: string;
  preview: string;
  open: string;
  properties: string;
  versionHistory: string;
  confirmDelete: string;
  folderName: string;
  newName: string;
  people: string;
  groups: string;
  links: string;
  view: string;
  edit: string;
  message: string;
  notify: string;
  createLink: string;
  copyLink: string;
  revoke: string;
  inherited: string;
  grantAccess: string;
  send: string;
  email: string;
  peopleSearchPlaceholder: string;
  searchingPeople: string;
  peopleSearchError: string;
  canView: string;
  canEdit: string;
  copyLinkSuccess: string;
  expiration: string;
  anyone: string;
  organization: string;
  specificPeople: string;
  destination: string;
  selectDestination: string;
  noSubfolders: string;
  invalidDestination: string;
  confirmDeleteMany: string;
  uploading: string;
  permissionDenied: string;
  notFound: string;
  conflict: string;
  throttled: string;
  authRequired: string;
  unknownError: string;
  selected: string;
  restore: string;
  noResults: string;
  refresh: string;
  selectAll: string;
  openInSharePoint: string;
  restoreSuccess: string;
  downloadVersion: string;
  uploadConflict: string;
  conflictRename: string;
  conflictReplace: string;
  conflictFail: string;
  cancelUpload: string;
  type: string;
  createdBy: string;
  modifiedBy: string;
  itemType: string;
  folder: string;
  file: string;
  loadMore: string;
  grantAccessSuccess: string;
  checkout: string;
  checkin: string;
  discardCheckout: string;
  checkinComment: string;
  createWord: string;
  createExcel: string;
  createPowerPoint: string;
  print: string;
  uploadFolder: string;
  updateExpiration: string;
  searchScopeFolder: string;
  searchScopeLibrary: string;
  filters: string;
  fileType: string;
  modifiedAfter: string;
  modifiedBefore: string;
  author: string;
  applyFilters: string;
  clearFilters: string;
  contentType: string;
  metadata: string;
  activity: string;
  activityLog: string;
  noActivity: string;
  deletedFrom: string;
  deletedDate: string;
  confirmDeletePermanent: string;
  confirmDeletePermanentMany: string;
  confirmEmptyRecycleBin: string;
  restoreSuccessRecycleBin: string;
  copyInProgress: string;
  copyComplete: string;
  copyFailed: string;
  bulkEditMetadata: string;
  bulkEditMetadataHint: string;
  leaveBlankToSkip: string;
  sensitivityLabel: string;
  dropToMove: string;
  operationSuccess: string;
  details: string;
  emptyHint: string;
  checkedOut: string;
  yes: string;
  no: string;
  itemCount: string;
  resultCount: string;
  dropFilesHere: string;
  columns: string;
  resetColumns: string;
  created: string;
  compact: string;
  createOrUpload: string;
  allDocuments: string;
  addColumn: string;
  searchLibraryPlaceholder: string;
  selectToViewDetails: string;
  uploadFiles: string;
  sortAscending: string;
  sortDescending: string;
  hideColumn: string;
  moreActions: string;
}

const vi: Messages = {
  files: "Tệp",
  search: "Tìm kiếm",
  searchPlaceholder: "Tìm tệp hoặc thư mục",
  newFolder: "Thư mục mới",
  upload: "Tải lên",
  list: "Danh sách",
  grid: "Lưới",
  name: "Tên",
  modified: "Sửa đổi",
  size: "Kích thước",
  empty: "Thư mục trống",
  loading: "Đang tải...",
  retry: "Thử lại",
  cancel: "Hủy",
  save: "Lưu",
  create: "Tạo",
  delete: "Xóa",
  rename: "Đổi tên",
  copy: "Sao chép",
  move: "Di chuyển",
  download: "Tải xuống",
  share: "Chia sẻ",
  manageAccess: "Quản lý quyền",
  preview: "Xem trước",
  open: "Mở",
  properties: "Thuộc tính",
  versionHistory: "Lịch sử phiên bản",
  confirmDelete: "Bạn có chắc muốn xóa mục này? Mục sẽ được đưa vào thùng rác SharePoint.",
  folderName: "Tên thư mục",
  newName: "Tên mới",
  people: "Người có quyền",
  groups: "Nhóm",
  links: "Liên kết",
  view: "Xem",
  edit: "Chỉnh sửa",
  message: "Lời nhắn",
  notify: "Gửi thông báo",
  createLink: "Tạo liên kết",
  copyLink: "Sao chép liên kết",
  revoke: "Thu hồi",
  inherited: "Kế thừa từ thư mục cha",
  grantAccess: "Cấp quyền",
  send: "Gửi",
  email: "Email",
  peopleSearchPlaceholder: "Thêm tên, nhóm hoặc email",
  searchingPeople: "Đang tìm người...",
  peopleSearchError: "Không thể tải danh sách người dùng. Vui lòng đăng xuất và đăng nhập lại.",
  canView: "Có thể xem",
  canEdit: "Có thể chỉnh sửa",
  copyLinkSuccess: "Đã sao chép liên kết",
  expiration: "Hết hạn",
  anyone: "Bất kỳ ai có liên kết",
  organization: "Mọi người trong tổ chức",
  specificPeople: "Người được chọn",
  destination: "Thư mục đích",
  selectDestination: "Chọn thư mục đích. Bấm vào thư mục để đi vào trong.",
  noSubfolders: "Không có thư mục con",
  invalidDestination: "Không thể sao chép/di chuyển vào chính mục đang chọn.",
  confirmDeleteMany: "Bạn có chắc muốn xóa các mục đã chọn? Chúng sẽ được đưa vào thùng rác SharePoint.",
  uploading: "Đang tải lên...",
  permissionDenied: "Bạn không có quyền thực hiện thao tác này.",
  notFound: "Tệp hoặc thư mục không còn tồn tại.",
  conflict: "Tên tệp/thư mục bị trùng.",
  throttled: "Quá nhiều yêu cầu, đang thử lại...",
  authRequired: "Cần đăng nhập lại Microsoft.",
  unknownError: "Đã xảy ra lỗi.",
  selected: "đã chọn",
  restore: "Khôi phục",
  noResults: "Không tìm thấy kết quả",
  refresh: "Làm mới",
  selectAll: "Chọn tất cả",
  openInSharePoint: "Mở trên SharePoint",
  restoreSuccess: "Đã khôi phục phiên bản",
  downloadVersion: "Tải phiên bản",
  uploadConflict: "Nếu trùng tên",
  conflictRename: "Đổi tên tự động",
  conflictReplace: "Thay thế",
  conflictFail: "Hủy tải lên",
  cancelUpload: "Hủy tải lên",
  type: "Loại",
  createdBy: "Người tạo",
  modifiedBy: "Người sửa",
  itemType: "Loại mục",
  folder: "Thư mục",
  file: "Tệp",
  loadMore: "Tải thêm",
  grantAccessSuccess: "Đã cấp quyền",
  checkout: "Check-out",
  checkin: "Check-in",
  discardCheckout: "Hủy check-out",
  checkinComment: "Ghi chú check-in",
  createWord: "Tài liệu Word",
  createExcel: "Bảng tính Excel",
  createPowerPoint: "Trình chiếu PowerPoint",
  print: "In",
  uploadFolder: "Tải thư mục lên",
  updateExpiration: "Cập nhật hết hạn",
  searchScopeFolder: "Thư mục hiện tại",
  searchScopeLibrary: "Toàn thư viện",
  filters: "Bộ lọc",
  fileType: "Loại tệp",
  modifiedAfter: "Sửa sau ngày",
  modifiedBefore: "Sửa trước ngày",
  author: "Tác giả",
  applyFilters: "Áp dụng",
  clearFilters: "Xóa lọc",
  contentType: "Loại nội dung",
  metadata: "Metadata",
  activity: "Hoạt động",
  activityLog: "Nhật ký hoạt động",
  noActivity: "Không có hoạt động",
  deletedFrom: "Vị trí gốc",
  deletedDate: "Ngày xóa",
  confirmDeletePermanent: "Xóa vĩnh viễn mục này? Không thể hoàn tác.",
  confirmDeletePermanentMany: "Xóa vĩnh viễn các mục đã chọn? Không thể hoàn tác.",
  confirmEmptyRecycleBin: "Xóa vĩnh viễn tất cả mục trong thùng rác hiện tại?",
  restoreSuccessRecycleBin: "Đã khôi phục mục",
  copyInProgress: "Đang sao chép...",
  copyComplete: "Sao chép hoàn tất",
  copyFailed: "Sao chép thất bại",
  bulkEditMetadata: "Sửa metadata hàng loạt",
  bulkEditMetadataHint: "Chỉ các trường có giá trị mới được cập nhật cho tất cả mục đã chọn.",
  leaveBlankToSkip: "Để trống nếu bỏ qua",
  sensitivityLabel: "Nhãn mật",
  dropToMove: "Thả để di chuyển",
  operationSuccess: "Thao tác thành công",
  details: "Chi tiết",
  emptyHint: "Kéo thả tệp vào đây hoặc dùng Tải lên / Thư mục mới",
  checkedOut: "Đang check-out",
  yes: "Có",
  no: "Không",
  itemCount: "mục",
  resultCount: "kết quả",
  dropFilesHere: "Thả tệp vào đây để tải lên",
  columns: "Cột",
  resetColumns: "Đặt lại",
  created: "Ngày tạo",
  compact: "Gọn",
  createOrUpload: "Tạo hoặc tải lên",
  allDocuments: "Tất cả tài liệu",
  addColumn: "Thêm cột",
  searchLibraryPlaceholder: "Tìm kiếm thư viện này",
  selectToViewDetails: "Chọn một tệp hoặc thư mục để xem chi tiết.",
  uploadFiles: "Tải tệp lên",
  sortAscending: "Sắp xếp tăng dần",
  sortDescending: "Sắp xếp giảm dần",
  hideColumn: "Ẩn cột",
  moreActions: "Thêm",
};

const en: Messages = {
  files: "Files",
  search: "Search",
  searchPlaceholder: "Search files or folders",
  newFolder: "New folder",
  upload: "Upload",
  list: "List",
  grid: "Grid",
  name: "Name",
  modified: "Modified",
  size: "Size",
  empty: "This folder is empty",
  loading: "Loading...",
  retry: "Retry",
  cancel: "Cancel",
  save: "Save",
  create: "Create",
  delete: "Delete",
  rename: "Rename",
  copy: "Copy",
  move: "Move",
  download: "Download",
  share: "Share",
  manageAccess: "Manage access",
  preview: "Preview",
  open: "Open",
  properties: "Properties",
  versionHistory: "Version history",
  confirmDelete: "Delete this item? It will go to the SharePoint recycle bin.",
  folderName: "Folder name",
  newName: "New name",
  people: "People with access",
  groups: "Groups",
  links: "Links",
  view: "View",
  edit: "Edit",
  message: "Message",
  notify: "Notify people",
  createLink: "Create link",
  copyLink: "Copy link",
  revoke: "Revoke",
  inherited: "Inherited from parent",
  grantAccess: "Grant access",
  send: "Send",
  email: "Email",
  peopleSearchPlaceholder: "Add a name, group, or email",
  searchingPeople: "Searching people...",
  peopleSearchError: "Could not load people suggestions. Please sign out and sign in again.",
  canView: "Can view",
  canEdit: "Can edit",
  copyLinkSuccess: "Link copied",
  expiration: "Expiration",
  anyone: "Anyone with the link",
  organization: "People in the organization",
  specificPeople: "Specific people",
  destination: "Destination folder",
  selectDestination: "Choose a destination folder. Click a folder to open it.",
  noSubfolders: "No subfolders",
  invalidDestination: "You cannot copy or move an item into itself.",
  confirmDeleteMany: "Delete the selected items? They will go to the SharePoint recycle bin.",
  uploading: "Uploading...",
  permissionDenied: "You do not have permission to do this.",
  notFound: "The file or folder no longer exists.",
  conflict: "A file or folder with this name already exists.",
  throttled: "Too many requests, retrying...",
  authRequired: "Microsoft sign-in is required.",
  unknownError: "Something went wrong.",
  selected: "selected",
  restore: "Restore",
  noResults: "No results",
  refresh: "Refresh",
  selectAll: "Select all",
  openInSharePoint: "Open in SharePoint",
  restoreSuccess: "Version restored",
  downloadVersion: "Download version",
  uploadConflict: "If name exists",
  conflictRename: "Rename automatically",
  conflictReplace: "Replace",
  conflictFail: "Cancel upload",
  cancelUpload: "Cancel upload",
  type: "Type",
  createdBy: "Created by",
  modifiedBy: "Modified by",
  itemType: "Item type",
  folder: "Folder",
  file: "File",
  loadMore: "Load more",
  grantAccessSuccess: "Access granted",
  checkout: "Check out",
  checkin: "Check in",
  discardCheckout: "Discard check-out",
  checkinComment: "Check-in comment",
  createWord: "Word document",
  createExcel: "Excel workbook",
  createPowerPoint: "PowerPoint presentation",
  print: "Print",
  uploadFolder: "Upload folder",
  updateExpiration: "Update expiration",
  searchScopeFolder: "Current folder",
  searchScopeLibrary: "Entire library",
  filters: "Filters",
  fileType: "File type",
  modifiedAfter: "Modified after",
  modifiedBefore: "Modified before",
  author: "Author",
  applyFilters: "Apply",
  clearFilters: "Clear",
  contentType: "Content type",
  metadata: "Metadata",
  activity: "Activity",
  activityLog: "Activity log",
  noActivity: "No activity recorded",
  deletedFrom: "Original location",
  deletedDate: "Deleted",
  confirmDeletePermanent: "Permanently delete this item? This cannot be undone.",
  confirmDeletePermanentMany: "Permanently delete the selected items? This cannot be undone.",
  confirmEmptyRecycleBin: "Permanently delete all items in the current recycle bin view?",
  restoreSuccessRecycleBin: "Item restored",
  copyInProgress: "Copying...",
  copyComplete: "Copy completed",
  copyFailed: "Copy failed",
  bulkEditMetadata: "Bulk edit metadata",
  bulkEditMetadataHint: "Only fields with values are applied to all selected items.",
  leaveBlankToSkip: "Leave blank to skip",
  sensitivityLabel: "Sensitivity label",
  dropToMove: "Drop to move here",
  operationSuccess: "Operation completed",
  details: "Details",
  emptyHint: "Drag files here or use Upload / New folder",
  checkedOut: "Checked out",
  yes: "Yes",
  no: "No",
  itemCount: "items",
  resultCount: "results",
  dropFilesHere: "Drop files here to upload",
  columns: "Columns",
  resetColumns: "Reset",
  created: "Created",
  compact: "Compact",
  createOrUpload: "Create or upload",
  allDocuments: "All documents",
  addColumn: "Add column",
  searchLibraryPlaceholder: "Search this library",
  selectToViewDetails: "Select a file or folder to see details.",
  uploadFiles: "Upload files",
  sortAscending: "Sort ascending",
  sortDescending: "Sort descending",
  hideColumn: "Hide column",
  moreActions: "More",
};

export function getMessages(locale: string): Messages {
  return locale.toLowerCase().startsWith("en") ? en : vi;
}
