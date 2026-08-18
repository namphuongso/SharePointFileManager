import { Skeleton, SkeletonItem } from "@fluentui/react-components";

export function LibrarySkeleton({ compact = false }: { compact?: boolean }) {
  return (
    <div className={`spm-library-skeleton ${compact ? "spm-density-compact" : ""}`} aria-busy="true">
      <div className="spm-skeleton-header">
        <SkeletonItem size={16} style={{ width: 16 }} />
        <SkeletonItem size={16} style={{ width: "38%" }} />
        <SkeletonItem size={16} style={{ width: "20%" }} />
        <SkeletonItem size={16} style={{ width: "18%" }} />
      </div>
      <Skeleton animation="wave">
        {Array.from({ length: compact ? 10 : 8 }, (_, index) => (
          <div className="spm-skeleton-row" key={index}>
            <SkeletonItem size={16} style={{ width: 16 }} />
            <SkeletonItem size={24} shape="square" />
            <SkeletonItem size={16} style={{ width: `${42 - (index % 3) * 6}%` }} />
            <SkeletonItem size={16} style={{ width: "18%" }} />
            <SkeletonItem size={16} style={{ width: "15%" }} />
          </div>
        ))}
      </Skeleton>
    </div>
  );
}
