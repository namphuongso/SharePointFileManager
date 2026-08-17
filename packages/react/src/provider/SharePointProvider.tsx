import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { SharePointClient, type SharePointConfig } from "@namphuongso/sharepoint-file-manager-core";
import { useMemo, useState, type ReactNode } from "react";
import { getMessages, type Messages } from "../i18n/messages";
import { SharePointContext } from "./context";

export interface SharePointProviderProps {
  config: SharePointConfig;
  locale?: string;
  messages?: Partial<Messages>;
  queryClient?: QueryClient;
  children: ReactNode;
}

export function SharePointProvider({
  config,
  locale = "vi-VN",
  messages,
  queryClient,
  children,
}: SharePointProviderProps) {
  const client = useMemo(() => new SharePointClient(config), [config]);
  const [fallbackClient] = useState(() => new QueryClient());
  const resolvedClient = queryClient ?? fallbackClient;
  const resolvedMessages = { ...getMessages(locale), ...messages };

  return (
    <QueryClientProvider client={resolvedClient}>
      <SharePointContext.Provider value={{ client, locale, messages: resolvedMessages }}>
        {children}
      </SharePointContext.Provider>
    </QueryClientProvider>
  );
}
