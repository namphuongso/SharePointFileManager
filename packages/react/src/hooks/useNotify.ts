import type { NotifyPayload } from "@namphuongso/sharepoint-file-manager-core";
import { useSharePoint } from "../provider/context";

export function useNotify() {
  const { onNotify } = useSharePoint();
  return (payload: NotifyPayload) => {
    onNotify?.(payload);
  };
}
