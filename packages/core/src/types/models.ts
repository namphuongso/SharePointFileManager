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
  lastModifiedDateTime?: string;
  /** Giá trị cột list (không $select). Key = InternalName / EntityPropertyName. */
  fields?: Record<string, unknown>;
}

/** Một trang list items (phân trang @odata.nextLink, không $skip). */
export interface FolderChildrenPage {
  items: SharePointItem[];
  nextLink?: string;
}

/**
 * Cột library (SP.Field) để option ẩn/hiện.
 * `internalName` là key ổn định; `title` là nhãn UI.
 */
export interface SharePointField {
  id?: string;
  title: string;
  internalName: string;
  entityPropertyName?: string;
  typeAsString?: string;
  fieldTypeKind?: number;
  hidden: boolean;
  readOnly: boolean;
  required: boolean;
  sortable?: boolean;
  filterable?: boolean;
}

export interface SharePointAppConfig {
  siteId?: string;
  siteUrl?: string;
  scopes?: string[];
  tokenProvider: TokenProvider;
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
