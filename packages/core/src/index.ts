export { defaultSharePointScopes, DEFAULT_SHAREPOINT_SCOPE_SUFFIX } from "./auth/token-provider";

export { SharePointClient } from "./client";
export { SharePointRestClient } from "./rest/client";
export { resolveConfig } from "./config/resolve-config";
export { createSharePointConfig } from "./config/create-sharepoint-config";
export { resolveAppConfig } from "./config/resolve-app-config";

export { SharePointError, SharePointErrorCode, isSharePointError } from "./errors/sharepoint-error";
export type { SharePointErrorOptions } from "./types/errors";
export { mapRestError } from "./errors/map-rest-error";
export { mapStatusToCode } from "./errors/map-status-to-code";

export { FolderService } from "./services/folder";
export { resolveLibrary } from "./services/library";

export type {
  TokenProvider,
  TokenRequest,
  SharePointConfig,
  SharePointAppConfig,
  ResolvedSharePointAppConfig,
  SharePointLibraryTarget,
  ResolvedSharePointConfig,
  SharePointItem,
  SharePointItemType,
  LibraryContext,
} from "./types/models";
