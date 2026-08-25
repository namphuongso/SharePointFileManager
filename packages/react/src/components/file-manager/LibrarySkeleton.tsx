import { Skeleton, SkeletonItem } from "@fluentui/react-components";

/** Placeholder khi đang tải listChildren. */

export function LibrarySkeleton() {
  return (
    <div style={{ padding: "8px 12px" }} aria-busy="true">
      <Skeleton>
        {Array.from({ length: 8 }, (_, index) => (
          <div
            key={index}
            style={{
              display: "grid",
              gridTemplateColumns: "20px minmax(0, 2fr) 120px 88px 160px",
              gap: 12,
              alignItems: "center",
              height: 42,
              borderBottom: "1px solid var(--colorNeutralStroke2)",
            }}
          >
            <SkeletonItem shape="square" size={20} />
            <SkeletonItem size={16} style={{ width: `${55 - (index % 4) * 8}%` }} />
            <SkeletonItem size={16} />
            <SkeletonItem size={16} />
            <SkeletonItem size={16} />
          </div>
        ))}
      </Skeleton>
    </div>
  );
}
