export type { TokenProvider, TokenRequest } from "./auth/token-provider";
export { defaultSharePointScopes, DEFAULT_SHAREPOINT_SCOPE_SUFFIX } from "./auth/token-provider";

export { SharePointClient } from "./client";
export { SharePointRestClient } from "./rest/client";
export { resolveConfig } from "./config/resolve-config";
export { createSharePointConfig } from "./config/create-sharepoint-config";

export { SharePointError, SharePointErrorCode, isSharePointError } from "./errors/sharepoint-error";
export { mapRestError, mapStatusToCode } from "./errors/map-rest-error";

export { FolderService } from "./services/folder";

export type {
  SharePointConfig,
  SharePointAppConfig,
  ResolvedSharePointAppConfig,
  SharePointLibraryTarget,
  ResolvedSharePointConfig,
  SharePointItem,
  SharePointItemType,
} from "./types/models";
