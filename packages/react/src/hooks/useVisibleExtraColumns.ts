import { useCallback, useEffect, useMemo, useState } from "react";
import {
  readVisibleExtraColumns,
  writeVisibleExtraColumns,
} from "../utils/visibleColumnsStorage";

/** Cột extra mặc định giống All Documents SharePoint: Người sửa đổi. */
const DEFAULT_EXTRA_VISIBLE = ["Editor"];

/**
 * Ẩn/hiện cột extra trên UI; persist localStorage theo site + thư viện.
 * Tick không gọi REST — $select vẫn lấy đủ catalog.
 */
export function useVisibleExtraColumns(scope: string, extraNames: readonly string[]) {
  const [visibleExtra, setVisibleExtra] = useState<Set<string> | undefined>(() => {
    const stored = readVisibleExtraColumns(scope);
    return stored ? new Set(stored) : undefined;
  });

  useEffect(() => {
    const stored = readVisibleExtraColumns(scope);
    setVisibleExtra(stored ? new Set(stored) : undefined);
  }, [scope]);

  const visible = useMemo(() => {
    const allowed = new Set(extraNames);
    if (visibleExtra) {
      return new Set([...visibleExtra].filter((name) => allowed.has(name)));
    }
    return new Set(DEFAULT_EXTRA_VISIBLE.filter((name) => allowed.has(name)));
  }, [visibleExtra, extraNames]);

  const setVisible = useCallback(
    (next: Set<string>) => {
      writeVisibleExtraColumns(scope, [...next]);
      setVisibleExtra(next);
    },
    [scope],
  );

  return { visible, setVisible };
}
