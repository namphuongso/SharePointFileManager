import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { FluentProvider } from "@fluentui/react-components";
import { SharePointClient, type SharePointConfig } from "@namphuongso/sharepoint-file-manager-core";
import { useMemo, useState, type ReactNode } from "react";
import { sharePointDarkTheme, sharePointLightTheme } from "../fluent/theme";
import { getMessages, type Messages } from "../i18n/messages";
import { SharePointContext } from "./context";

export interface SharePointProviderProps {
  config: SharePointConfig;
  locale?: string;
  messages?: Partial<Messages>;
  queryClient?: QueryClient;
  children: ReactNode;
  theme?: "light" | "dark" | "system";
}

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
  const isDark =
    theme === "dark" ||
    (theme === "system" &&
      typeof matchMedia !== "undefined" &&
      matchMedia("(prefers-color-scheme: dark)").matches);

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
