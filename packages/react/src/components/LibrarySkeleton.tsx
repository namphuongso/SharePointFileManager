import { Skeleton, SkeletonItem } from "@fluentui/react-components";

export function LibrarySkeleton() {
  return (
    <div style={{ padding: 12 }} aria-busy="true">
      <Skeleton>
        {Array.from({ length: 8 }, (_, index) => (
          <div
            key={index}
            style={{
              display: "grid",
              gridTemplateColumns: "24px 1fr 120px 80px",
              gap: 12,
              alignItems: "center",
              marginBottom: 10,
            }}
          >
            <SkeletonItem shape="square" size={24} />
            <SkeletonItem size={16} style={{ width: `${55 - (index % 4) * 8}%` }} />
            <SkeletonItem size={16} />
            <SkeletonItem size={16} />
          </div>
        ))}
      </Skeleton>
    </div>
  );
}
