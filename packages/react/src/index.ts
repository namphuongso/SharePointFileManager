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
export { useFolderChildren, getErrorMessage } from "./hooks/hooks";

export {
  SharePointClient,
  SharePointError,
  SharePointErrorCode,
  defaultSharePointScopes,
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
} from "@namphuongso/sharepoint-file-manager-core";
