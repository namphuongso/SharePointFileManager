export { createMsalTokenProvider } from "./auth/createMsalTokenProvider";
export type { MsalSilentTokenSource } from "./auth/createMsalTokenProvider";
export { SharePointProvider } from "./provider/SharePointProvider";
export { useSharePoint } from "./provider/context";
export { SharePointFileManager } from "./components/SharePointFileManager";
export type { SharePointFileManagerProps } from "./components/SharePointFileManager";
export {
  useSharePointSite,
  useSiteDrives,
  useSiteLists,
  findDriveByName,
  findListByName,
  useFolderChildren as useFiles,
  useFolderChildren as useFolders,
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
  getErrorMessage,
} from "./hooks/hooks";

export {
  SharePointClient,
  SharePointError,
  SharePointErrorCode,
  DEFAULT_GRAPH_SCOPES,
  isSharePointError,
} from "@namphuongso/sharepoint-file-manager-core";

export type {
  TokenProvider,
  TokenRequest,
  SharePointConfig,
  SharePointItem,
  SharePointPermission,
  FeatureConfig,
  UploadProgress,
  DriveInfo,
  SiteInfo,
  SharePointListInfo,
} from "@namphuongso/sharepoint-file-manager-core";
