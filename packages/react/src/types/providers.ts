import type { ReactNode } from "react";
import type { QueryClient } from "@tanstack/react-query";
import type {
  ResolvedSharePointAppConfig,
  SharePointAppConfig,
  SharePointClient,
  SharePointConfig,
  SharePointLibraryTarget,
} from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "./messages";

export interface SharePointProviderProps {
  config: SharePointConfig;
  locale?: string;
  messages?: Partial<Messages>;
  queryClient?: QueryClient;
  children: ReactNode;
  theme?: "light" | "dark" | "system";
}

export interface SharePointContextValue {
  client: SharePointClient;
  locale: string;
  messages: Messages;
}

export interface SharePointAppProviderProps {
  config: SharePointAppConfig;
  locale?: string;
  messages?: Partial<Messages>;
  children: ReactNode;
}

export type SharePointAppStatus = "ready" | "error";

export interface SharePointAppContextValue {
  appConfig: ResolvedSharePointAppConfig | null;
  status: SharePointAppStatus;
  error: unknown;
  locale?: string;
  messages?: Partial<Messages>;
  createConfig: (target?: SharePointLibraryTarget) => SharePointConfig;
}
