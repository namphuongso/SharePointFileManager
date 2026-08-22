import { useQuery } from "@tanstack/react-query";
import { useSharePoint } from "../provider/context";
import { queryKeys } from "./queryKeys";

/** Một cấp con của folder hiện tại. queryKey gắn site + thư viện + folderId. */
export function useFolderChildren(folderId: string | undefined) {
  const { client } = useSharePoint();
  return useQuery({
    queryKey: queryKeys.children(client.config.siteId, client.cacheScope, folderId ?? ""),
    enabled: Boolean(folderId),
    queryFn: ({ signal }) => client.folders.listChildren(folderId!, { signal }),
  });
}
