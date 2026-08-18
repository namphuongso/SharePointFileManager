export interface UserInfo {
  id?: string;
  displayName?: string;
  email?: string;
}

export type DirectoryPersonKind = "user" | "group" | "email";

/** Person or group from Microsoft Graph people/directory search. */
export interface DirectoryPerson {
  key: string;
  displayName: string;
  email?: string;
  objectId?: string;
  kind: DirectoryPersonKind;
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
  /** Checkout / lock hints from Graph publication or listItem fields. */
  capabilities?: Partial<ItemCapabilities>;
  listItemId?: string;
  contentType?: string;
  /** Custom SharePoint list column values keyed by internal column name. */
  metadata?: Record<string, string | number | boolean | null>;
  sensitivityLabel?: string;
}

export interface ItemCapabilities {
  canRead?: boolean;
  canRename: boolean;
  canDelete: boolean;
  canDownload: boolean;
  canMove: boolean;
  canCopy: boolean;
  canShare: boolean;
  canCheckout: boolean;
  canCheckin: boolean;
  canDiscardCheckout: boolean;
  canManagePermissions?: boolean;
  canViewVersions?: boolean;
  canRestore?: boolean;
  canPreview?: boolean;
  isCheckedOut: boolean;
  checkedOutByMe: boolean;
  checkedOutBy?: UserInfo;
}

export type OfficeFileKind = "word" | "excel" | "powerpoint";

export type SortField = "name" | "modified" | "size" | "created";
export type SortDirection = "asc" | "desc";

export interface PagedResult<T> {
  items: T[];
  nextLink?: string;
  hasMore?: boolean;
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
  chunkSize?: number;
  resume?: { uploadUrl: string; nextOffset?: number };
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
  onCopyProgress?: (progress: CopyOperationProgress) => void;
}

export type CopyOperationPhase = "starting" | "monitoring" | "completed" | "failed";

export interface CopyOperationProgress {
  phase: CopyOperationPhase;
  /** 0–100 while monitoring async copy job. */
  percent?: number;
}

export type SearchScope = "folder" | "library";

export interface SearchFilters {
  fileType?: string;
  modifiedAfter?: string;
  modifiedBefore?: string;
  author?: string;
}

export interface SearchOptions {
  query: string;
  scope?: SearchScope;
  folderId?: string;
  filters?: SearchFilters;
  top?: number;
  nextLink?: string;
  /** Offset for Microsoft Search API library queries. */
  from?: number;
  signal?: AbortSignal;
}

export interface ListColumn {
  id: string;
  name: string;
  displayName: string;
  readOnly: boolean;
  hidden: boolean;
  type?: string;
}

export interface ListItemFields {
  itemId: string;
  listItemId?: string;
  contentType?: string;
  fields: Record<string, string | number | boolean | null>;
}

export interface DriveItemActivity {
  id: string;
  action?: string;
  actor?: UserInfo;
  timestamp?: string;
  description?: string;
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
  properties?: boolean;
  checkout?: boolean;
  createOfficeFile?: boolean;
  globalSearch?: boolean;
  metadata?: boolean;
  activityLog?: boolean;
  infiniteScroll?: boolean;
  dragDropMove?: boolean;
  bulkMetadata?: boolean;
  copyProgress?: boolean;
  enableDeltaSync?: boolean;
  enableAnalytics?: boolean;
  enableActivities?: boolean;
}

export type NotifyType = "success" | "error" | "info";

export interface NotifyPayload {
  type: NotifyType;
  message: string;
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
