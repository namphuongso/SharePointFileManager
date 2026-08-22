import type {
  ResolvedSharePointAppConfig,
  SharePointAppConfig,
  SharePointConfig,
  SharePointLibraryTarget,
} from "@namphuongso/sharepoint-file-manager-core";
import { createSharePointConfig } from "@namphuongso/sharepoint-file-manager-core";
import { createContext, useContext, useMemo, type ReactNode } from "react";
import type { Messages } from "../i18n/messages";

export interface SharePointAppProviderProps {
  config: SharePointAppConfig;
  locale?: string;
  messages?: Partial<Messages>;
  children: ReactNode;
}

export type SharePointAppStatus = "ready" | "error";

interface SharePointAppContextValue {
  appConfig: ResolvedSharePointAppConfig | null;
  status: SharePointAppStatus;
  error: unknown;
  locale?: string;
  messages?: Partial<Messages>;
  createConfig: (target?: SharePointLibraryTarget) => SharePointConfig;
}

const SharePointAppContext = createContext<SharePointAppContextValue | null>(null);

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

export function useSharePointApp(): SharePointAppContextValue {
  const value = useContext(SharePointAppContext);
  if (!value) {
    throw new Error("useSharePointApp must be used within SharePointAppProvider");
  }
  return value;
}

export function useOptionalSharePointApp(): SharePointAppContextValue | null {
  return useContext(SharePointAppContext);
}
