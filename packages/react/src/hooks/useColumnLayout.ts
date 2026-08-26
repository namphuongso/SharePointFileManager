import { useCallback, useEffect, useMemo, useState } from "react";
import {
  defaultColumnWidth,
  mergeColumnOrder,
  moveColumn,
  persistColumnOrder,
  type ColumnLayout,
} from "../utils/columnLayout";
import { readColumnLayout, writeColumnLayout } from "../utils/columnLayoutStorage";

/**
 * Thứ tự + độ rộng cột trên UI; persist localStorage theo site + thư viện.
 * Kéo thả / resize không gọi REST.
 */
export function useColumnLayout(scope: string, columnIds: readonly string[]) {
  const [layout, setLayout] = useState<ColumnLayout | undefined>(() => readColumnLayout(scope));

  useEffect(() => {
    setLayout(readColumnLayout(scope));
  }, [scope]);

  const order = useMemo(() => mergeColumnOrder(layout?.order, columnIds), [layout?.order, columnIds]);

  const widths = useMemo(() => {
    const next: Record<string, number> = {};
    for (const id of order) {
      const stored = layout?.widths[id];
      next[id] = typeof stored === "number" && stored > 0 ? stored : defaultColumnWidth(id);
    }
    return next;
  }, [layout, order]);

  const onReorder = useCallback(
    (fromField: string, toField: string, place: "before" | "after") => {
      setLayout((current) => {
        const visible = mergeColumnOrder(current?.order, columnIds);
        const nextVisible = moveColumn(visible, fromField, toField, place);
        const next: ColumnLayout = {
          order: persistColumnOrder(current?.order, nextVisible),
          widths: current?.widths ?? {},
        };
        writeColumnLayout(scope, next);
        return next;
      });
    },
    [columnIds, scope],
  );

  const onResize = useCallback((field: string, width: number) => {
    setLayout((current) => ({
      order: current?.order ?? [...columnIds],
      widths: { ...(current?.widths ?? {}), [field]: width },
    }));
  }, [columnIds]);

  const onResizeEnd = useCallback(
    (field: string, width: number) => {
      setLayout((current) => {
        const next: ColumnLayout = {
          order: persistColumnOrder(current?.order, mergeColumnOrder(current?.order, columnIds)),
          widths: { ...(current?.widths ?? {}), [field]: width },
        };
        writeColumnLayout(scope, next);
        return next;
      });
    },
    [columnIds, scope],
  );

  return { order, widths, onReorder, onResize, onResizeEnd };
}
