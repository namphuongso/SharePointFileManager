import { useQuery } from "@tanstack/react-query";
import { useSharePoint } from "../provider/context";
import { queryKeys } from "./queryKeys";

export function useFolderChildren(folderId: string | undefined) {
  const { client } = useSharePoint();
  return useQuery({
    queryKey: queryKeys.children(client.config.siteId, client.cacheScope, folderId ?? ""),
    enabled: Boolean(folderId),
    queryFn: ({ signal }) => client.folders.listChildren(folderId!, { signal }),
  });
}

export function getErrorMessage(error: unknown, fallback = "Unknown error"): string {
  if (!error) return fallback;
  if (typeof error === "object" && error !== null && "message" in error) {
    const message = (error as { message?: unknown }).message;
    if (typeof message === "string" && message.trim()) return message;
  }
  return fallback;
}
