export { createMsalTokenProvider } from "./auth/createMsalTokenProvider";
export type { MsalSilentTokenSource } from "./auth/createMsalTokenProvider";
export { SharePointProvider } from "./provider/SharePointProvider";
export {
  SharePointAppProvider,
  useSharePointApp,
  useOptionalSharePointApp,
} from "./provider/SharePointAppProvider";
export type { SharePointAppProviderProps } from "./provider/SharePointAppProvider";
export { sharePointLightTheme, sharePointDarkTheme } from "./fluent/theme";
export { useSharePoint } from "./provider/context";
export { SharePointFileManager } from "./components/SharePointFileManager";
export type { SharePointFileManagerProps } from "./components/SharePointFileManager";
export {
  SharePointShell,
  type SharePointShellHeaderConfig,
  type SharePointShellNavigationItem,
} from "./components/SharePointShell";
export {
  useSharePointSite,
  useSiteDrives,
  useSiteLists,
  findDriveByName,
  findListByName,
  useFolderChildren,
  useItem,
  useSearchItems as useSearch,
  usePermissions,
  useVersions,
  useCreateFolder,
  useRenameItem,
  useDeleteItem,
  useCopyItem,
  useMoveItem,
  useUploadFile as useUpload,
  useDownloadFile as useDownload,
  useInvite,
  useCreateLink,
  useRemovePermission,
  usePeopleSearch,
  useUpdatePermission,
  useRestoreVersion,
  useDownloadVersion,
  useCheckout,
  useFolderChildrenInfinite,
  useListColumns,
  useListItemFields,
  useUpdateListItemFields,
  useBulkUpdateListItemFields,
  useItemActivities,
  decodeLibraryPage,
  getErrorMessage,
} from "./hooks/hooks";
export { useNotify } from "./hooks/useNotify";

export {
  SharePointClient,
  SharePointError,
  SharePointErrorCode,
  DEFAULT_GRAPH_SCOPES,
  createSharePointConfig,
  isSharePointError,
} from "@namphuongso/sharepoint-file-manager-core";

export type {
  TokenProvider,
  TokenRequest,
  SharePointConfig,
  SharePointAppConfig,
  ResolvedSharePointAppConfig,
  SharePointLibraryTarget,
  SharePointItem,
  SharePointPermission,
  FeatureConfig,
  UploadProgress,
  DriveInfo,
  SiteInfo,
  SharePointListInfo,
  DirectoryPerson,
  NotifyPayload,
  CopyOperationProgress,
  CreateLinkScope,
  ShareScope,
  ItemCapabilities,
} from "@namphuongso/sharepoint-file-manager-core";
