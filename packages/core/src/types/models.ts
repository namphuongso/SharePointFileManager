export type SharePointItemType = "file" | "folder";

export interface TokenRequest {
  scopes: string[];
  forceRefresh?: boolean;
}

/** App host cung cấp. Thư viện không đăng nhập hộ và không lưu token. */
export interface TokenProvider {
  getAccessToken(request: TokenRequest): Promise<string>;
}

export interface SharePointItem {
  id: string;
  name: string;
  type: SharePointItemType;
  size?: number;
  /** Folder: tổng số item con trực tiếp (file + folder) từ Folder/ItemCount. */
  childItemCount?: number;
  lastModifiedDateTime?: string;
  /** Giá trị cột do SharePoint trả về. Key = InternalName. */
  fields?: Record<string, unknown>;
}

/** Một trang list items (phân trang @odata.nextLink, không $skip). */
export interface FolderChildrenPage {
  items: SharePointItem[];
  nextLink?: string;
}

/**
 * Cột option ẩn/hiện lấy từ SharePoint.
 * `internalName` khớp `$select` items; `title` là nhãn SharePoint trả theo locale.
 */
export interface SharePointField {
  title: string;
  internalName: string;
  /** TypeAsString từ SharePoint — dùng để tránh select computed field. */
  typeAsString?: string;
}

export interface SharePointAppConfig {
  siteId?: string;
  siteUrl?: string;
  scopes?: string[];
  tokenProvider: TokenProvider;

  /** Ngôn ngữ nhãn cột — gửi Accept-Language khi GET cột (vd. vi-VN, en-US). */
  locale?: string;
}

export type ResolvedSharePointAppConfig = SharePointAppConfig & {
  siteId: string;
  siteUrl: string;
};

export interface SharePointLibraryTarget {
  libraryName?: string;
  listId?: string;
  rootItemId?: string;
}

export interface SharePointConfig extends SharePointAppConfig, SharePointLibraryTarget {}

export interface ResolvedSharePointConfig extends SharePointConfig {
  siteId: string;
  siteUrl: string;
  rootItemId: string;
  scopes: string[];
}

/** List đang dùng và folder gốc. FolderService list theo path; UniqueId để nhận biết root. */
export interface LibraryContext {
  listId: string;
  listTitle: string;
  rootFolderServerRelativeUrl: string;
  rootFolderUniqueId: string;
  entityTypeName?: string;
}

/** Tùy chọn POST tạo thư mục con. */
export interface CreateFolderOptions {
  signal?: AbortSignal;
}

/** Tùy chọn POST soft-delete (recycle) file/folder. */
export interface DeleteItemOptions {
  signal?: AbortSignal;
}

/** Tùy chọn POST upload file vào folder. */
export interface UploadFileOptions {
  /** Tên file trên SharePoint — mặc định File.name nếu content là File. */
  fileName?: string;
  /** Ghi đè khi trùng tên — mặc định false. */
  overwrite?: boolean;
  signal?: AbortSignal;
}
