import { createContext, useContext } from "react";
import type { SharePointAppContextValue } from "../types";

/** Hook đọc cấu hình app (siteUrl + token). Optional = không throw khi ngoài provider. */

export const SharePointAppContext = createContext<SharePointAppContextValue | null>(null);

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
