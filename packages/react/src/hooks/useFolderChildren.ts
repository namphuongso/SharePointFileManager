import { useInfiniteQuery } from "@tanstack/react-query";
import { useSharePoint } from "../provider/context";
import { queryKeys } from "./queryKeys";

/** Một cấp con; phân trang theo @odata.nextLink. */
export function useFolderChildren(folderId: string | undefined) {
  const { client } = useSharePoint();
  return useInfiniteQuery({
    queryKey: queryKeys.children(client.config.siteId, client.cacheScope, folderId ?? ""),
    enabled: Boolean(folderId),
    initialPageParam: undefined as string | undefined,
    queryFn: ({ signal, pageParam }) =>
      client.folders.listChildren(folderId!, { signal, nextLink: pageParam }),
    getNextPageParam: (page) => page.nextLink,
  });
}
