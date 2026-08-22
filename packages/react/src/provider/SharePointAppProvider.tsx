import { createSharePointConfig, resolveAppConfig } from "@namphuongso/sharepoint-file-manager-core";
import { useMemo } from "react";
import type {
  SharePointAppContextValue,
  SharePointAppProviderProps,
  SharePointAppStatus,
} from "../types";
import { SharePointAppContext } from "./useSharePointApp";

/**
 * Cấp app: siteUrl + token. createConfig gắn libraryName cho từng trang.
 */
export function SharePointAppProvider({
  config,
  locale,
  messages,
  children,
}: SharePointAppProviderProps) {
  const resolvedAppConfig = useMemo(() => resolveAppConfig(config), [config]);
  const status: SharePointAppStatus = resolvedAppConfig ? "ready" : "error";
  const error = resolvedAppConfig
    ? undefined
    : new Error("SharePointAppProvider requires config.siteUrl");

  const value = useMemo<SharePointAppContextValue>(
    () => ({
      appConfig: resolvedAppConfig,
      status,
      error,
      locale,
      messages,
      createConfig: (target = {}) => {
        if (!resolvedAppConfig) {
          throw new Error("SharePoint app config is not ready yet");
        }
        return createSharePointConfig(resolvedAppConfig, target);
      },
    }),
    [resolvedAppConfig, status, error, locale, messages],
  );

  return <SharePointAppContext.Provider value={value}>{children}</SharePointAppContext.Provider>;
}
