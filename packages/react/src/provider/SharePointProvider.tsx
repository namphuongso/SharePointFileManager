import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import {
  FluentProvider,
  Toast,
  Toaster,
  ToastTitle,
  useId,
  useToastController,
} from "@fluentui/react-components";
import { SharePointClient, type NotifyPayload, type SharePointConfig } from "@namphuongso/sharepoint-file-manager-core";
import { useMemo, useState, type ReactNode } from "react";
import { sharePointDarkTheme, sharePointLightTheme } from "../fluent/theme";
import { getMessages, type Messages } from "../i18n/messages";
import { SharePointContext } from "./context";

export interface SharePointProviderProps {
  config: SharePointConfig;
  locale?: string;
  messages?: Partial<Messages>;
  queryClient?: QueryClient;
  onNotify?: (payload: NotifyPayload) => void;
  children: ReactNode;
  theme?: "light" | "dark" | "system";
  density?: "compact" | "comfortable";
  embedded?: boolean;
}

export function SharePointProvider({
  config,
  locale = "vi-VN",
  messages,
  queryClient,
  onNotify,
  children,
  theme = "light",
  density = "comfortable",
  embedded = true,
}: SharePointProviderProps) {
  const client = useMemo(() => new SharePointClient(config), [config]);
  const [fallbackClient] = useState(() => new QueryClient());
  const resolvedClient = queryClient ?? fallbackClient;
  const resolvedMessages = { ...getMessages(locale), ...messages };

  return (
    <QueryClientProvider client={resolvedClient}>
      <FluentProvider
        theme={theme === "dark" || (theme === "system" && typeof matchMedia !== "undefined" && matchMedia("(prefers-color-scheme: dark)").matches) ? sharePointDarkTheme : sharePointLightTheme}
        data-spm-density={density}
        data-spm-embedded={embedded ? "true" : "false"}
        style={{
          display: "flex",
          flexDirection: "column",
          height: "100%",
          minHeight: 0,
          backgroundColor: "transparent",
        }}
      >
        <NotificationBridge
          client={client}
          locale={locale}
          messages={resolvedMessages}
          onNotify={onNotify}
        >
          {children}
        </NotificationBridge>
      </FluentProvider>
    </QueryClientProvider>
  );
}

function NotificationBridge({
  client,
  locale,
  messages,
  onNotify,
  children,
}: {
  client: SharePointClient;
  locale: string;
  messages: Messages;
  onNotify?: (payload: NotifyPayload) => void;
  children: ReactNode;
}) {
  const toasterId = useId("spm-toaster");
  const { dispatchToast } = useToastController(toasterId);

  function handleNotify(payload: NotifyPayload) {
    dispatchToast(
      <Toast>
        <ToastTitle>{payload.message}</ToastTitle>
      </Toast>,
      { intent: payload.type === "info" ? "info" : payload.type, position: "top-end" },
    );
    onNotify?.(payload);
  }

  return (
    <SharePointContext.Provider value={{ client, locale, messages, onNotify: handleNotify }}>
      {children}
      <Toaster toasterId={toasterId} limit={4} pauseOnHover />
    </SharePointContext.Provider>
  );
}
