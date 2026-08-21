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
 * Domain model for UI. `capabilities` only carries checkout state Graph actually returns;
 * host FeatureConfig gates CRUD/share actions.
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
  /** Browser URL that opens like SharePoint web (Doc.aspx for Office files). */
  openUrl?: string;
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

/** Graph-derived checkout state only (v1.0 publication / listItem fields). */
export interface ItemCapabilities {
  isCheckedOut: boolean;
  checkedOutBy?: UserInfo;
}

export type SortField = "name" | "modified" | "size" | "created";
export type SortDirection = "asc" | "desc";

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

/** Scope on returned permission links (may include `users`). */
export type ShareScope = "anonymous" | "organization" | "users";

/** Scopes accepted by Graph v1.0 `createLink`. Specific people use `invite` instead. */
export type CreateLinkScope = "anonymous" | "organization";

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
  scope: CreateLinkScope;
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
  globalSearch?: boolean;
  metadata?: boolean;
  activityLog?: boolean;
  infiniteScroll?: boolean;
  dragDropMove?: boolean;
  bulkMetadata?: boolean;
  copyProgress?: boolean;
}

export type NotifyType = "success" | "error" | "info";

export interface NotifyPayload {
  type: NotifyType;
  message: string;
}

/**
 * App-level SharePoint settings configured once by the host (site + auth + features).
 * Pass `siteId` or `siteUrl` (`SharePointAppProvider` resolves URL → id).
 * Per-route/module usage merges a library target via `createSharePointConfig`.
 */
export interface SharePointAppConfig {
  /** Graph site id. Optional when `siteUrl` is set — resolved inside `SharePointAppProvider`. */
  siteId?: string;
  /** SharePoint web URL, e.g. https://contoso.sharepoint.com/sites/eOffice */
  siteUrl?: string;
  scopes?: string[];
  graphBaseUrl?: string;
  tokenProvider: import("../auth/token-provider").TokenProvider;
  features?: FeatureConfig;
}

/** App config after siteId is known (ready to merge with a library target). */
export type ResolvedSharePointAppConfig = SharePointAppConfig & { siteId: string };

/** Per-feature/route targeting: which library (and optional folder) to show. */
export interface SharePointLibraryTarget {
  /** Resolve drive by live Graph name, e.g. "eDocumentTest" or "Documents". */
  libraryName?: string;
  /** Resolve drive via GET /sites/{siteId}/lists/{listId}/drive. */
  listId?: string;
  driveId?: string;
  /** Starting folder drive item id; default "root". */
  rootItemId?: string;
}

export interface SharePointConfig extends SharePointAppConfig, SharePointLibraryTarget {}

export interface ResolvedSharePointConfig extends SharePointConfig {
  siteId: string;
  rootItemId: string;
  scopes: string[];
  graphBaseUrl: string;
  features: Required<FeatureConfig>;
}
