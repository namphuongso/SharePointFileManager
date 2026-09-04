import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { FluentProvider } from "@fluentui/react-components";
import { SharePointClient } from "@namphuongso/sharepoint-file-manager-core";
import { useCallback, useMemo, useRef, useState } from "react";
import { isDarkTheme } from "../fluent/isDarkTheme";
import { sharePointDarkTheme, sharePointLightTheme } from "../fluent/theme";
import { getMessages } from "../i18n/messages";
import { ToastViewport } from "../notify";
import type { NotifyApi } from "../notify";
import type { SharePointProviderProps } from "../types";
import { SharePointContext } from "./context";

const LOCALE_STORAGE_KEY = "sp_file_manager_locale";

/** Locale dùng chung giữa AppProvider và FileManager; ưu tiên lựa chọn của người dùng. */
export function resolveInitialLocale(configLocale?: string, propLocale?: string): string {
  if (typeof window !== "undefined") {
    const persisted = window.localStorage.getItem(LOCALE_STORAGE_KEY);
    if (persisted) return persisted;
  }
  return configLocale || propLocale || "vi-VN";
}

/** Đọc locale người dùng đã chọn; F5 vẫn giữ đúng ngôn ngữ trên toolbar. */
function readPersistedLocale(): string | undefined {
  if (typeof window === "undefined") return undefined;
  return window.localStorage.getItem(LOCALE_STORAGE_KEY) ?? undefined;
}

/** Khóa tái tạo client khi đổi site/thư viện hoặc instance thiếu service (HMR / bản cũ). */
function clientInstanceKey(config: SharePointProviderProps["config"], locale: string): string {
  return [config.siteUrl, config.libraryName ?? "", config.listId ?? "", config.rootItemId ?? "", locale].join(
    "|",
  );
}

/**
 * Cấp trang: tạo SharePointClient, QueryClient, Fluent theme, i18n.
 */
export function SharePointProvider({
  config,
  locale = "vi-VN",
  messages,
  queryClient,
  children,
  theme = "light",
}: SharePointProviderProps) {
  const initialLocale = resolveInitialLocale(config.locale, locale);
  const instanceKey = clientInstanceKey(config, initialLocale);

  const [clientState, setClientState] = useState(() => ({
    key: instanceKey,
    client: new SharePointClient({ ...config, locale: initialLocale }),
  }));
  // HMR / host giữ instance cũ: tạo lại khi thiếu search hoặc đổi target.
  if (
    clientState.key !== instanceKey ||
    typeof clientState.client.search?.listAccessible !== "function"
  ) {
    setClientState({
      key: instanceKey,
      client: new SharePointClient({ ...config, locale: initialLocale }),
    });
  }
  const client = clientState.client;

  const [fallbackClient] = useState(() => new QueryClient());
  const [runtimeLocale, setRuntimeLocale] = useState(() => ({
    source: initialLocale,
    value: initialLocale,
  }));
  const resolvedClient = queryClient ?? fallbackClient;
  const requestLocale = initialLocale;
  const activeLocale = runtimeLocale.source === requestLocale ? runtimeLocale.value : requestLocale;

  if (runtimeLocale.source !== requestLocale) {
    setRuntimeLocale({ source: requestLocale, value: requestLocale });
  }

  if (client.config.locale !== activeLocale) {
    client.setLocale(activeLocale);
  }

  const resolvedMessages = getMessages(activeLocale, messages);
  const isDark = isDarkTheme(theme);

  /**
   * API toast — `ToastViewport` gọi `onReady` khi controller sẵn sàng.
   * Giữ trong ref để dispatch an toàn kể cả khi hook bắn toast đồng bộ trong commit đầu
   * (trước effect mount Toaster). Context trả cùng một object qua useMemo.
   */
  const notifyRef = useRef<NotifyApi>({
    success: () => {},
    info: () => "",
    progress: () => "",
    error: () => {},
    update: () => {},
    dismiss: () => {},
  });
  const notify = useMemo<NotifyApi>(
    () => ({
      success: (title, subtitle) => notifyRef.current.success(title, subtitle),
      info: (title, subtitle) => notifyRef.current.info(title, subtitle),
      progress: (title, subtitle) => notifyRef.current.progress(title, subtitle),
      error: (title, subtitle) => notifyRef.current.error(title, subtitle),
      update: (id, options) => notifyRef.current.update(id, options),
      dismiss: (id) => notifyRef.current.dismiss(id),
    }),
    [],
  );
  const handleToastReady = useCallback((api: NotifyApi) => {
    notifyRef.current = api;
  }, []);

  return (
    <QueryClientProvider client={resolvedClient}>
      <FluentProvider
        theme={isDark ? sharePointDarkTheme : sharePointLightTheme}
        style={{
          display: "flex",
          flexDirection: "column",
          height: "100%",
          minHeight: 0,
          backgroundColor: "transparent",
        }}
      >
        <SharePointContext.Provider
          value={{
            client,
            locale: activeLocale,
            setLocale: (locale) => {
              if (typeof window !== "undefined") {
                window.localStorage.setItem(LOCALE_STORAGE_KEY, locale);
              }
              setRuntimeLocale({ source: locale, value: locale });
            },
            messages: resolvedMessages,
            notify,
          }}
        >
          {children}
          <ToastViewport onReady={handleToastReady} />
        </SharePointContext.Provider>
      </FluentProvider>
    </QueryClientProvider>
  );
}
