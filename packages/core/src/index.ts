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

export { FolderService, isSortableLibraryField, isSortableSearchField } from "./services/folder";
export { FieldService, FIXED_LIBRARY_FIELD_NAMES } from "./services/fields";
export { FileService } from "./services/files";
export { PermissionService, hasPermissionKind, toItemCapabilities } from "./services/permissions";
export { SearchService } from "./services/search";
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
  SharePointField,
  FolderChildrenPage,
  LibraryContext,
} from "./types/models";
export type { ListChildrenSort, ListSortDirection } from "./types/rest";
export { PermissionKind } from "./types/permissions";
export type {
  EffectiveBasePermissionsDto,
  ItemCapabilities,
  PermissionItemType,
} from "./types/permissions";
export type {
  SearchAccessibleOptions,
  SearchAccessiblePage,
} from "./types/search";
