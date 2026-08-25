import { useEffect, useRef } from "react";

/**
 * Infinite scroll: khi sentinel vào vùng cuộn của root thì gọi loadMore.
 * Root phải là phần tử overflow (listPane), không phải viewport cửa sổ.
 */
export function useLoadMoreOnScroll(
  enabled: boolean,
  isLoading: boolean,
  loadMore: () => void,
) {
  const rootRef = useRef<HTMLDivElement>(null);
  const sentinelRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const root = rootRef.current;
    const sentinel = sentinelRef.current;
    if (!enabled || !root || !sentinel) return;

    const observer = new IntersectionObserver(
      (entries) => {
        if (!entries[0]?.isIntersecting || isLoading) return;
        loadMore();
      },
      { root, rootMargin: "120px 0px", threshold: 0 },
    );
    observer.observe(sentinel);
    return () => observer.disconnect();
  }, [enabled, isLoading, loadMore]);

  return { rootRef, sentinelRef };
}
