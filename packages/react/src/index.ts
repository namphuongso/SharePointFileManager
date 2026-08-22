export { createMsalTokenProvider } from "./auth/createMsalTokenProvider";
export { SharePointProvider } from "./provider/SharePointProvider";
export { SharePointAppProvider } from "./provider/SharePointAppProvider";
export { useSharePointApp, useOptionalSharePointApp } from "./provider/useSharePointApp";
export { sharePointLightTheme, sharePointDarkTheme } from "./fluent/theme";
export { useSharePoint } from "./provider/context";
export { SharePointFileManager } from "./components/file-manager/SharePointFileManager";
export { useFolderChildren } from "./hooks/useFolderChildren";
export { getErrorMessage } from "./hooks/getErrorMessage";

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

export type {
  Messages,
  MsalSilentTokenSource,
  SharePointAppProviderProps,
  SharePointFileManagerProps,
  SharePointProviderProps,
} from "./types";
