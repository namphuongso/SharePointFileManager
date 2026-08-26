import { useInfiniteQuery } from "@tanstack/react-query";
import type { ListChildrenSort } from "@namphuongso/sharepoint-file-manager-core";
import { useSharePoint } from "../provider/context";
import { queryKeys } from "./queryKeys";

/** Một cấp con; phân trang theo @odata.nextLink. Sort nằm queryKey để GET lại trang đầu. */
export function useFolderChildren(
  folderId: string | undefined,
  sort?: ListChildrenSort,
) {
  const { client, locale } = useSharePoint();
  return useInfiniteQuery({
    queryKey: queryKeys.children(
      client.config.siteId,
      client.cacheScope,
      client.config.locale ?? locale,
      folderId ?? "",
      sort?.field ?? "",
      sort?.direction ?? "",
    ),
    enabled: Boolean(folderId),
    initialPageParam: undefined as string | undefined,
    queryFn: ({ signal, pageParam }) =>
      client.folders.listChildren(folderId!, {
        signal,
        nextLink: pageParam,
        sort: pageParam ? undefined : sort,
      }),
    getNextPageParam: (page) => page.nextLink,
  });
}
