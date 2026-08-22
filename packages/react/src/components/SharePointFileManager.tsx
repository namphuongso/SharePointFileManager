import { useMemo, useState } from "react";
import type { SharePointConfig, SharePointItem, SharePointLibraryTarget } from "@namphuongso/sharepoint-file-manager-core";
import {
  Breadcrumb,
  BreadcrumbButton,
  BreadcrumbDivider,
  BreadcrumbItem,
  Button,
} from "@fluentui/react-components";
import { ArrowClockwiseRegular, HomeRegular } from "@fluentui/react-icons";
import { SharePointProvider } from "../provider/SharePointProvider";
import { useOptionalSharePointApp } from "../provider/SharePointAppProvider";
import { useSharePoint } from "../provider/context";
import { getErrorMessage, useFolderChildren } from "../hooks/hooks";
import type { Messages } from "../i18n/messages";
import { EmptyState } from "./EmptyState";
import { FileList } from "./FileView";
import { LibrarySkeleton } from "./LibrarySkeleton";
import { ErrorBanner } from "./ui";

export interface SharePointFileManagerProps extends SharePointLibraryTarget {
  config?: SharePointConfig;
  locale?: string;
  className?: string;
  title?: string;
  messages?: Partial<Messages>;
  theme?: "light" | "dark" | "system";
}

export function SharePointFileManager(props: SharePointFileManagerProps) {
  const app = useOptionalSharePointApp();
  const target: SharePointLibraryTarget = {
    libraryName: props.libraryName,
    listId: props.listId,
    rootItemId: props.rootItemId,
  };
  const hasTarget = Boolean(target.libraryName || target.listId || target.rootItemId);

  if (!props.config && app && app.status === "error") {
    return (
      <p style={{ padding: 16, color: "#616161", fontSize: 14 }}>
        Không tải được cấu hình SharePoint. Kiểm tra `siteUrl` và quyền SharePoint.
      </p>
    );
  }

  const config =
    props.config ?? (app?.status === "ready" && app.appConfig ? app.createConfig(target) : undefined);

  if (!config) {
    throw new Error(
      "SharePointFileManager requires `config`, or `libraryName`/`listId` inside SharePointAppProvider",
    );
  }

  return (
    <SharePointProvider
      config={hasTarget && props.config ? { ...props.config, ...target } : config}
      locale={props.locale ?? app?.locale}
      messages={props.messages ?? app?.messages}
      theme={props.theme}
    >
      <FileBrowser className={props.className} title={props.title} />
    </SharePointProvider>
  );
}

interface Crumb {
  id: string;
  name: string;
}

function FileBrowser({ className, title }: { className?: string; title?: string }) {
  const { client, locale, messages } = useSharePoint();
  const rootId = client.config.rootItemId;
  const [crumbs, setCrumbs] = useState<Crumb[]>([{ id: rootId, name: title ?? messages.files }]);
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
