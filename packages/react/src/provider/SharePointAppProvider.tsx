import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import type {
  NotifyPayload,
  ResolvedSharePointAppConfig,
  SharePointAppConfig,
  SharePointConfig,
  SharePointLibraryTarget,
} from "@namphuongso/sharepoint-file-manager-core";
import { createSharePointConfig } from "@namphuongso/sharepoint-file-manager-core";
import { createContext, useContext, useMemo, useState, type ReactNode } from "react";
import { useSharePointSite } from "../hooks/hooks";
import type { Messages } from "../i18n/messages";

export interface SharePointAppProviderProps {
  /**
   * Site + token + scopes + default features — configure once at app root.
   * Use `siteUrl` or `siteId`; when only `siteUrl` is set, the provider resolves Graph site id.
   */
  config: SharePointAppConfig;
  locale?: string;
  messages?: Partial<Messages>;
  onNotify?: (payload: NotifyPayload) => void;
  children: ReactNode;
}

export type SharePointAppStatus = "loading" | "ready" | "error";

interface SharePointAppContextValue {
  appConfig: ResolvedSharePointAppConfig | null;
  status: SharePointAppStatus;
  error: unknown;
  locale?: string;
  messages?: Partial<Messages>;
  onNotify?: (payload: NotifyPayload) => void;
  /** Merge a per-route library target onto the resolved app config. */
  createConfig: (target?: SharePointLibraryTarget) => SharePointConfig;
}

const SharePointAppContext = createContext<SharePointAppContextValue | null>(null);

function SharePointAppProviderInner({
  config,
  locale,
  messages,
  onNotify,
  children,
}: SharePointAppProviderProps) {
  const needsResolve = Boolean(config.siteUrl?.trim() && !config.siteId?.trim());

  const siteQuery = useSharePointSite(
    needsResolve
      ? {
          tokenProvider: config.tokenProvider,
          siteUrl: config.siteUrl!.trim(),
          scopes: config.scopes,
          graphBaseUrl: config.graphBaseUrl,
        }
      : undefined,
  );

  const siteId = (config.siteId?.trim() || siteQuery.data?.id || "").trim() || undefined;

  const status: SharePointAppStatus = !siteId
    ? needsResolve && siteQuery.isLoading
      ? "loading"
      : siteQuery.isError
        ? "error"
        : "loading"
    : "ready";

  const resolvedAppConfig = useMemo<ResolvedSharePointAppConfig | null>(() => {
    if (!siteId) return null;
    return { ...config, siteId };
  }, [config, siteId]);

  const value = useMemo<SharePointAppContextValue>(
    () => ({
      appConfig: resolvedAppConfig,
      status,
      error: siteQuery.error,
      locale,
      messages,
      onNotify,
      createConfig: (target = {}) => {
        if (!resolvedAppConfig) {
          throw new Error("SharePoint app config is not ready yet");
        }
        return createSharePointConfig(resolvedAppConfig, target);
      },
    }),
    [resolvedAppConfig, status, siteQuery.error, locale, messages, onNotify],
  );

  return <SharePointAppContext.Provider value={value}>{children}</SharePointAppContext.Provider>;
}

/**
 * Host app root: pass site + token once. Feature routes only use `libraryName`.
 *
 * ```tsx
 * <SharePointAppProvider config={{ siteUrl, tokenProvider, features }}>
 *   <App />
 * </SharePointAppProvider>
 *
 * // route
 * <SharePointFileManager libraryName="eDocumentTest" />
 * ```
 */
export function SharePointAppProvider(props: SharePointAppProviderProps) {
  const [queryClient] = useState(() => new QueryClient());

  return (
    <QueryClientProvider client={queryClient}>
      <SharePointAppProviderInner {...props} />
    </QueryClientProvider>
  );
}

export function useSharePointApp(): SharePointAppContextValue {
  const value = useContext(SharePointAppContext);
  if (!value) {
    throw new Error("useSharePointApp must be used inside SharePointAppProvider");
  }
  return value;
}

export function useOptionalSharePointApp(): SharePointAppContextValue | null {
  return useContext(SharePointAppContext);
}
