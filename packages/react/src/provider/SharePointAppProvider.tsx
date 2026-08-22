import type { ResolvedSharePointAppConfig } from "@namphuongso/sharepoint-file-manager-core";
import { createSharePointConfig } from "@namphuongso/sharepoint-file-manager-core";
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
  const siteUrl = config.siteUrl?.trim() ?? "";
  const status: SharePointAppStatus = siteUrl ? "ready" : "error";
  const error = siteUrl ? undefined : new Error("SharePointAppProvider requires config.siteUrl");

  const resolvedAppConfig = useMemo<ResolvedSharePointAppConfig | null>(() => {
    if (!siteUrl) return null;
    return {
      ...config,
      siteUrl,
      siteId: config.siteId?.trim() || siteUrl,
    };
  }, [config, siteUrl]);

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
