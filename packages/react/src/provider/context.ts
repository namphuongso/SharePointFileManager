import { createContext, useContext } from "react";
import type { SharePointContextValue } from "../types";

/** Context cấp trang: client REST + locale + chuỗi UI. */

export const SharePointContext = createContext<SharePointContextValue | null>(null);

export function useSharePoint(): SharePointContextValue {
  const value = useContext(SharePointContext);
  if (!value) {
    throw new Error("useSharePoint must be used inside SharePointProvider");
  }
  return value;
}
