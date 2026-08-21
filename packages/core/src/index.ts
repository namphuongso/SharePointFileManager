export type { TokenProvider, TokenRequest } from "./auth/token-provider";
export { DEFAULT_GRAPH_SCOPES } from "./auth/token-provider";

export { SharePointClient } from "./client";
export { GraphClient } from "./graph/client";
export { resolveConfig, DEFAULT_FEATURES } from "./config/resolve-config";
export { createSharePointConfig } from "./config/create-sharepoint-config";

export { SharePointError, SharePointErrorCode, isSharePointError } from "./errors/sharepoint-error";
export { mapGraphError, mapStatusToCode } from "./errors/map-graph-error";
export { mapDriveItem } from "./mappers/item";
export { canPerformItemAction, isOfficeOnlineFile } from "./utils/item-actions";
export type { ItemAction } from "./utils/item-actions";
export {
  buildSharePointDocOpenUrl,
  isDirectFileDownloadUrl,
  isOfficeFileName,
  resolveGraphItemOpenUrl,
  resolveItemOpenUrl,
} from "./utils/sharepoint-open-url";
export type { GraphSharePointIds } from "./utils/sharepoint-open-url";
export { mapPermission } from "./mappers/permission";
export { mapGraphPerson, mapGraphUser, mapGraphGroup, mapTypedEmail, toInviteRecipient } from "./mappers/person";

export { DriveService, findDriveByName, mapGraphDrive } from "./services/drive";
export { SiteService, siteIdentifierFromUrl, findListByName, isVisibleSharePointList } from "./services/site";
export { FileService } from "./services/file";
export { FolderService } from "./services/folder";
export { UploadService } from "./services/upload";
export { SharingService, PermissionService } from "./services/sharing";
export { SearchService, buildSearchKql, decodeLibraryPage } from "./services/search";
export { PeopleService } from "./services/people";
export { CheckoutService } from "./services/checkout";
export { ListItemService } from "./services/list-item";
export { ActivityService } from "./services/activity";
export { isVisibleListColumn, mapListColumn } from "./mappers/list-item";
export { itemsVisibleInFolder } from "./utils/list-visible-items";

export type {
  SharePointConfig,
  SharePointAppConfig,
  ResolvedSharePointAppConfig,
  SharePointLibraryTarget,
  ResolvedSharePointConfig,
  FeatureConfig,
  SharePointItem,
  SharePointItemType,
  PagedResult,
  DriveInfo,
  SiteInfo,
  SharePointListInfo,
  UserInfo,
  DirectoryPerson,
  DirectoryPersonKind,
  ConflictBehavior,
  ShareRole,
  ShareLinkType,
  ShareScope,
  CreateLinkScope,
  PermissionKind,
  ShareLinkInfo,
  SharePointPermission,
  FileVersion,
  PreviewInfo,
  InviteRecipient,
  InviteOptions,
  CreateLinkOptions,
  UploadOptions,
  UploadProgress,
  CopyMoveOptions,
  CopyOperationProgress,
  CopyOperationPhase,
  NotifyPayload,
  NotifyType,
  SearchOptions,
  SearchScope,
  SearchFilters,
  ListColumn,
  ListItemFields,
  DriveItemActivity,
  ItemCapabilities,
  SortField,
  SortDirection,
} from "./types/models";
