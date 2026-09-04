import { useFileManagerStyles } from "./useFileManagerStyles";

/** Số dòng skeleton mặc định. */
const DEFAULT_ROW_COUNT = 8;
/** Khoảng cách giữa các dòng (ms) — tạo cảm giác "đang nạp" tuần tự. */
const STAGGER_STEP_MS = 50;

/**
 * Placeholder khi đang tải listChildren / search.
 * Mỗi dòng fade-in lệch nhau (stagger) + shimmer gradient chạy ngang vô tận.
 * Layout 5 cột khớp với bảng: icon | tên | sửa đổi | cột option | size.
 */
export function LibrarySkeleton({ rowCount = DEFAULT_ROW_COUNT }: { rowCount?: number } = {}) {
  const styles = useFileManagerStyles();
  return (
    <div className={styles.skeletonWrap} aria-busy="true" aria-live="polite">
      {Array.from({ length: rowCount }, (_, index) => (
        <div
          key={index}
          className={styles.skeletonRow}
          style={{ animationDelay: `${index * STAGGER_STEP_MS}ms` }}
        >
          <div className={styles.skeletonShimmer} style={{ width: 20, height: 20 }} />
          <div
            className={styles.skeletonShimmer}
            style={{ width: `${55 - (index % 4) * 8}%`, height: 14 }}
          />
          <div className={styles.skeletonShimmer} style={{ width: 88, height: 14 }} />
          <div className={styles.skeletonShimmer} style={{ width: 72, height: 14 }} />
          <div className={styles.skeletonShimmer} style={{ width: 56, height: 14 }} />
        </div>
      ))}
    </div>
  );
}
