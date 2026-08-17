export type { TokenProvider, TokenRequest } from "./auth/token-provider";
export { DEFAULT_GRAPH_SCOPES } from "./auth/token-provider";

export { SharePointClient } from "./client";
export { GraphClient } from "./graph/client";
export { resolveConfig, DEFAULT_FEATURES, isFeatureEnabled } from "./config/resolve-config";

export { SharePointError, SharePointErrorCode, isSharePointError } from "./errors/sharepoint-error";
export { mapGraphError, mapStatusToCode } from "./errors/map-graph-error";
export { mapDriveItem } from "./mappers/item";
export { mapPermission } from "./mappers/permission";

export { DriveService, findDriveByName, mapGraphDrive } from "./services/drive";
export { SiteService, siteIdentifierFromUrl, findListByName, isVisibleSharePointList } from "./services/site";
export { FileService } from "./services/file";
export { FolderService } from "./services/folder";
export { UploadService } from "./services/upload";
export { SharingService, PermissionService } from "./services/sharing";
export { SearchService } from "./services/search";

export type {
  SharePointConfig,
  ResolvedSharePointConfig,
  FeatureConfig,
  SharePointItem,
  SharePointItemType,
  PagedResult,
  DriveInfo,
  SiteInfo,
  SharePointListInfo,
  UserInfo,
  ConflictBehavior,
  ShareRole,
  ShareLinkType,
  ShareScope,
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
  SearchOptions,
} from "./types/models";
