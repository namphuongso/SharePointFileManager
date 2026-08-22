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
