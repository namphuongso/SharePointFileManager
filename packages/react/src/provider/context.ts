import { createContext, useContext } from "react";
import type { SharePointClient, NotifyPayload } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../i18n/messages";

export interface SharePointContextValue {
  client: SharePointClient;
  locale: string;
  messages: Messages;
  onNotify?: (payload: NotifyPayload) => void;
}

export const SharePointContext = createContext<SharePointContextValue | null>(null);

export function useSharePoint(): SharePointContextValue {
  const value = useContext(SharePointContext);
  if (!value) {
    throw new Error("useSharePoint must be used inside SharePointProvider");
  }
  return value;
}
