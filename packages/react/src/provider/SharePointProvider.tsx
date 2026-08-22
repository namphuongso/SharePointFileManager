import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { FluentProvider } from "@fluentui/react-components";
import { SharePointClient } from "@namphuongso/sharepoint-file-manager-core";
import { useMemo, useState } from "react";
import { isDarkTheme } from "../fluent/isDarkTheme";
import { sharePointDarkTheme, sharePointLightTheme } from "../fluent/theme";
import { getMessages } from "../i18n/messages";
import type { SharePointProviderProps } from "../types";
import { SharePointContext } from "./context";

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
  const client = useMemo(() => new SharePointClient(config), [config]);
  const [fallbackClient] = useState(() => new QueryClient());
  const resolvedClient = queryClient ?? fallbackClient;
  const resolvedMessages = { ...getMessages(locale), ...messages };
  const isDark = isDarkTheme(theme);

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
        <SharePointContext.Provider value={{ client, locale, messages: resolvedMessages }}>
          {children}
        </SharePointContext.Provider>
      </FluentProvider>
    </QueryClientProvider>
  );
}
