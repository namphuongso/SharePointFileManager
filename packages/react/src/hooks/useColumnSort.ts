import { useCallback, useEffect, useState } from "react";
import type {
  ListChildrenSort,
  ListSortDirection,
} from "@namphuongso/sharepoint-file-manager-core";
import { readColumnSort, writeColumnSort } from "../utils/columnSortStorage";

/**
 * Sort bảng → $orderby trang đầu. Persist localStorage theo site + thư viện.
 * Cùng hướng lần nữa thì về mặc định.
 */
export function useColumnSort(scope: string) {
  const [sort, setSort] = useState<ListChildrenSort | undefined>(() => readColumnSort(scope));

  useEffect(() => {
    setSort(readColumnSort(scope));
  }, [scope]);

  const onSort = useCallback(
    (field: string, direction: ListSortDirection) => {
      setSort((current) => {
        const next =
          current?.field === field && current.direction === direction
            ? undefined
            : { field, direction };
        writeColumnSort(scope, next);
        return next;
      });
    },
    [scope],
  );

  return { sort, onSort };
}
