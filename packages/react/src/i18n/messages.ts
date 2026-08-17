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
  email: string;
  expiration: string;
  anyone: string;
  organization: string;
  specificPeople: string;
  destination: string;
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
  email: "Email",
  expiration: "Hết hạn",
  anyone: "Bất kỳ ai có liên kết",
  organization: "Mọi người trong tổ chức",
  specificPeople: "Người được chọn",
  destination: "Thư mục đích",
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
  email: "Email",
  expiration: "Expiration",
  anyone: "Anyone with the link",
  organization: "People in the organization",
  specificPeople: "Specific people",
  destination: "Destination folder",
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
};

export function getMessages(locale: string): Messages {
  return locale.toLowerCase().startsWith("en") ? en : vi;
}
