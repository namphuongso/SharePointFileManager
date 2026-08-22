export type SharePointItemType = "file" | "folder";

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
  tokenProvider: import("../auth/token-provider").TokenProvider;
}

export type ResolvedSharePointAppConfig = SharePointAppConfig & { siteId: string; siteUrl: string };

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
