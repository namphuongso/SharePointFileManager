import { useMemo, useState } from "react";
import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import {
  Breadcrumb,
  BreadcrumbButton,
  BreadcrumbDivider,
  BreadcrumbItem,
  Button,
} from "@fluentui/react-components";
import { ArrowClockwiseRegular, HomeRegular } from "@fluentui/react-icons";
import { useSharePoint } from "../../provider/context";
import { getErrorMessage } from "../../hooks/getErrorMessage";
import { useFolderChildren } from "../../hooks/useFolderChildren";
import type { BreadcrumbCrumb, FileBrowserProps } from "../../types";
import { EmptyState } from "./EmptyState";
import { ErrorBanner } from "./ErrorBanner";
import { FileList } from "./FileView";
import { LibrarySkeleton } from "./LibrarySkeleton";

/**
 * Duyệt một thư viện: breadcrumb + list children.
 * Không tạo client — phải nằm trong SharePointProvider.
 */
export function FileBrowser({ className, title }: FileBrowserProps) {
  const { client, locale, messages } = useSharePoint();
  const rootId = client.config.rootItemId;
  const [crumbs, setCrumbs] = useState<BreadcrumbCrumb[]>([
    { id: rootId, name: title ?? messages.files },
  ]);
  const currentFolderId = crumbs[crumbs.length - 1]?.id ?? rootId;
  const childrenQuery = useFolderChildren(currentFolderId);
  const items = useMemo(() => childrenQuery.data ?? [], [childrenQuery.data]);

  function openFolder(item: SharePointItem) {
    if (item.type !== "folder") return;
    setCrumbs((prev) => [...prev, { id: item.id, name: item.name }]);
  }

  return (
    <div
      className={className}
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100%",
        minHeight: 0,
        background: "var(--colorNeutralBackground1)",
      }}
    >
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          padding: "8px 12px",
          borderBottom: "1px solid var(--colorNeutralStroke2)",
        }}
      >
        <Breadcrumb>
          {crumbs.map((crumb, index) => (
            <BreadcrumbItem key={crumb.id}>
              {index > 0 ? <BreadcrumbDivider /> : null}
              <BreadcrumbButton
                current={index === crumbs.length - 1}
                onClick={() => setCrumbs((prev) => prev.slice(0, index + 1))}
                icon={index === 0 ? <HomeRegular /> : undefined}
              >
                {crumb.name}
              </BreadcrumbButton>
            </BreadcrumbItem>
          ))}
        </Breadcrumb>
        <Button
          appearance="subtle"
          icon={<ArrowClockwiseRegular />}
          onClick={() => void childrenQuery.refetch()}
          aria-label={messages.refresh}
          title={messages.refresh}
        />
      </div>

      {childrenQuery.error ? (
        <ErrorBanner
          message={getErrorMessage(childrenQuery.error, messages.unknownError)}
          onRetry={() => void childrenQuery.refetch()}
          retryLabel={messages.retry}
        />
      ) : null}

      <div style={{ flex: 1, minHeight: 0, overflow: "auto" }}>
        {childrenQuery.isLoading ? <LibrarySkeleton /> : null}
        {!childrenQuery.isLoading && items.length === 0 ? <EmptyState messages={messages} /> : null}
        {!childrenQuery.isLoading && items.length > 0 ? (
          <FileList items={items} locale={locale} messages={messages} onOpenFolder={openFolder} />
        ) : null}
      </div>
    </div>
  );
}
