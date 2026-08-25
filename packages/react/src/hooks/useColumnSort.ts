import { useCallback, useState } from "react";
import type {
  ListChildrenSort,
  ListSortDirection,
} from "@namphuongso/sharepoint-file-manager-core";

/** Sort bảng → $orderby trang đầu. Cùng hướng lần nữa thì về mặc định. */
export function useColumnSort() {
  const [sort, setSort] = useState<ListChildrenSort | undefined>();
  const onSort = useCallback((field: string, direction: ListSortDirection) => {
    setSort((current) => {
      if (current?.field === field && current.direction === direction) return undefined;
      return { field, direction };
    });
  }, []);
  return { sort, onSort };
}
