export interface UserInfo {
  id?: string;
  displayName?: string;
  email?: string;
}

export type SharePointItemType = "file" | "folder";

/**
 * Domain model for UI. Values like canEdit/canDelete are hints from Graph,
 * not an application permission database.
 */
export interface SharePointItem {
  id: string;
  name: string;
  type: SharePointItemType;
  size?: number;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  createdBy?: UserInfo;
  lastModifiedBy?: UserInfo;
  webUrl?: string;
  mimeType?: string;
  parentId?: string;
  driveId?: string;
  eTag?: string;
  childCount?: number;
  downloadUrl?: string;
  thumbnailUrl?: string;
  canPreview?: boolean;
}

export interface PagedResult<T> {
  items: T[];
  nextLink?: string;
}

export interface DriveInfo {
  id: string;
  name?: string;
  webUrl?: string;
  driveType?: string;
}

export interface SiteInfo {
  id: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
}

export interface SharePointListInfo {
  id: string;
  name: string;
  displayName: string;
  webUrl?: string;
  template?: string;
  hasDrive: boolean;
}

export type ConflictBehavior = "fail" | "replace" | "rename";

export type ShareRole = "read" | "write";

export type ShareLinkType = "view" | "edit";

export type ShareScope = "anonymous" | "organization" | "users";

export type PermissionKind = "user" | "group" | "link" | "siteGroup" | "unknown";

export interface ShareLinkInfo {
  type?: ShareLinkType | string;
  scope?: ShareScope | string;
  webUrl?: string;
  preventsDownload?: boolean;
}

export interface SharePointPermission {
  id: string;
  roles: string[];
  kind: PermissionKind;
  inherited: boolean;
  inheritedFromId?: string;
  grantedTo?: UserInfo;
  grantedToGroup?: { id?: string; displayName?: string };
  link?: ShareLinkInfo;
  expirationDateTime?: string;
  /** Heuristic: inherited permissions usually cannot be deleted on the child item. */
  canRemove: boolean;
}

export interface FileVersion {
  id: string;
  lastModifiedDateTime?: string;
  lastModifiedBy?: UserInfo;
  size?: number;
}

export interface PreviewInfo {
  getUrl?: string;
  postUrl?: string;
  postParameters?: string;
}

export interface InviteRecipient {
  email?: string;
  objectId?: string;
}

export interface InviteOptions {
  recipients: InviteRecipient[];
  role: ShareRole;
  message?: string;
  sendInvitation?: boolean;
  requireSignIn?: boolean;
}

export interface CreateLinkOptions {
  type: ShareLinkType;
  scope: ShareScope;
  expirationDateTime?: string;
}

export interface UploadOptions {
  parentId: string;
  fileName: string;
  content: Blob | ArrayBuffer | Uint8Array;
  conflictBehavior?: ConflictBehavior;
  onProgress?: (progress: UploadProgress) => void;
  signal?: AbortSignal;
}

export interface UploadProgress {
  bytesUploaded: number;
  totalBytes: number;
  percent: number;
}

export interface CopyMoveOptions {
  itemId: string;
  destinationParentId: string;
  newName?: string;
  signal?: AbortSignal;
}

export interface SearchOptions {
  query: string;
  folderId?: string;
  top?: number;
  signal?: AbortSignal;
}

export interface FeatureConfig {
  upload?: boolean;
  download?: boolean;
  createFolder?: boolean;
  rename?: boolean;
  delete?: boolean;
  copy?: boolean;
  move?: boolean;
  share?: boolean;
  manageAccess?: boolean;
  search?: boolean;
  preview?: boolean;
  versionHistory?: boolean;
  openInSharePoint?: boolean;
}

export interface SharePointConfig {
  siteId: string;
  driveId?: string;
  /** Resolve drive via GET /sites/{siteId}/lists/{listId}/drive. */
  listId?: string;
  /** Resolve drive by live Graph name, e.g. "Documents" or "ISO Documents". */
  libraryName?: string;
  rootItemId?: string;
  scopes?: string[];
  graphBaseUrl?: string;
  tokenProvider: import("../auth/token-provider").TokenProvider;
  features?: FeatureConfig;
}

export interface ResolvedSharePointConfig extends SharePointConfig {
  rootItemId: string;
  scopes: string[];
  graphBaseUrl: string;
  features: Required<FeatureConfig>;
}
