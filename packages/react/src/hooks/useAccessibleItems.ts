import { useInfiniteQuery } from "@tanstack/react-query";
import type { ListChildrenSort } from "@namphuongso/sharepoint-file-manager-core";
import { useSharePoint } from "../provider/context";
import { queryKeys } from "./queryKeys";

interface UseAccessibleItemsOptions {
  enabled?: boolean;
  /** Cột option đang hiện — đưa vào selectproperties Search. */
  fieldInternalNames?: readonly string[];
}

/**
 * Flat list file/folder user được xem trong thư viện (Search REST).
 * Sort / cột nằm queryKey — đổi thì GET lại từ StartRow 0.
 */
export function useAccessibleItems(
  sort?: ListChildrenSort,
  options: UseAccessibleItemsOptions = {},
) {
  const { client } = useSharePoint();
  const fieldKey = (options.fieldInternalNames ?? []).join(",");
  return useInfiniteQuery({
    queryKey: queryKeys.accessible(
      client.config.siteId,
      client.cacheScope,
      sort?.field ?? "",
      sort?.direction ?? "",
      fieldKey,
    ),
    enabled: options.enabled ?? true,
    initialPageParam: 0,
    queryFn: ({ signal, pageParam }) => {
      const search = client.search;
      if (!search?.listAccessible) {
        throw new Error(
          "SharePointClient.search chưa sẵn sàng — hard refresh hoặc rebuild package core.",
        );
      }
      return search.listAccessible({
        signal,
        startRow: pageParam,
        sort,
        fieldInternalNames: options.fieldInternalNames,
      });
    },
    getNextPageParam: (page) => page.nextStartRow,
  });
}
