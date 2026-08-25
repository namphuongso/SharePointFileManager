import { Fragment, useCallback, useMemo, useState } from "react";
import { FIXED_LIBRARY_FIELD_NAMES, type SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import {
  Breadcrumb,
  BreadcrumbButton,
  BreadcrumbDivider,
  BreadcrumbItem,
  Button,
  mergeClasses,
} from "@fluentui/react-components";
import { ArrowClockwiseRegular, HomeRegular } from "@fluentui/react-icons";
import { useSharePoint } from "../../provider/context";
import { getErrorMessage } from "../../hooks/getErrorMessage";
import { useColumnSort } from "../../hooks/useColumnSort";
import { useFolderChildren } from "../../hooks/useFolderChildren";
import { useLoadMoreOnScroll } from "../../hooks/useLoadMoreOnScroll";
import { useLibraryFields } from "../../hooks/useLibraryFields";
import { useVisibleExtraColumns } from "../../hooks/useVisibleExtraColumns";
import { fieldLabel } from "../../i18n/messages";
import type { BreadcrumbCrumb, FileBrowserProps } from "../../types";
import { ColumnPicker } from "./ColumnPicker";
import { EmptyState } from "./EmptyState";
import { ErrorBanner } from "./ErrorBanner";
import { FileList } from "./FileView";
import { LanguageSwitcher } from "./LanguageSwitcher";
import { LibrarySkeleton } from "./LibrarySkeleton";
import { useFileManagerStyles } from "./useFileManagerStyles";

/** ItemChildCount: hiện Folder.ItemCount. FolderChildCount không có số tách — không đưa picker. */
const ITEM_CHILD_COUNT = "ItemChildCount";
const HIDE_FROM_PICKER = new Set([
  "FolderChildCount",
  "ComplianceAssetId",
  "DocIcon",
  "FileSize",
]);

/**
 * Duyệt một thư viện: command bar + list children (phân trang nextLink).
 * Không tạo client — phải nằm trong SharePointProvider.
 */
export function FileBrowser({ className, title, showLanguageSwitcher = true }: FileBrowserProps) {
  const styles = useFileManagerStyles();
  const { client, locale, messages } = useSharePoint();
  const rootId = client.config.rootItemId;
  const [crumbs, setCrumbs] = useState<BreadcrumbCrumb[]>([
    { id: rootId, name: title ?? messages.files },
  ]);
  const currentFolderId = crumbs[crumbs.length - 1]?.id ?? rootId;
  const columnScope = `${client.config.siteId}:${client.cacheScope}`;
  const { sort, onSort } = useColumnSort(columnScope);
  const childrenQuery = useFolderChildren(currentFolderId, sort);
  const fieldsQuery = useLibraryFields();
  const libraryFields = fieldsQuery.data ?? [];
  const selectableLibraryFields = useMemo(
    () =>
      libraryFields.filter(
        (field) =>
          !HIDE_FROM_PICKER.has(field.internalName) &&
          (field.typeAsString?.toLowerCase() !== "computed" ||
            field.internalName === ITEM_CHILD_COUNT),
      ),
    [libraryFields],
  );

  const extraNames = useMemo(
    () =>
      selectableLibraryFields
        .filter((f) => !FIXED_LIBRARY_FIELD_NAMES.has(f.internalName))
        .map((f) => f.internalName),
    [selectableLibraryFields],
  );
  const { visible: resolvedVisible, setVisible: setVisibleExtra } = useVisibleExtraColumns(
    columnScope,
    extraNames,
  );

  const items = useMemo(
    () => childrenQuery.data?.pages.flatMap((page) => page.items) ?? [],
    [childrenQuery.data],
  );
  const fetchNextPage = childrenQuery.fetchNextPage;
  const loadMore = useCallback(() => {
    void fetchNextPage();
  }, [fetchNextPage]);
  const { rootRef, sentinelRef } = useLoadMoreOnScroll(
    Boolean(childrenQuery.hasNextPage),
    childrenQuery.isFetchingNextPage,
    loadMore,
  );

  const extraColumns = useMemo(
    () =>
      selectableLibraryFields
        .filter(
          (f) =>
            resolvedVisible.has(f.internalName) &&
            !FIXED_LIBRARY_FIELD_NAMES.has(f.internalName),
        )
        .map((f) => ({
          internalName: f.internalName,
          title: fieldLabel(messages, f.internalName, f.title),
          typeAsString: f.typeAsString,
        })),
    [selectableLibraryFields, messages, resolvedVisible],
  );

  /** Nhãn 3 cột cố định chuẩn SharePoint theo locale; view thiếu cột thì fallback messages. */
  const fixedTitles = useMemo(() => {
    const titleBy = new Map(
      libraryFields.map((field) => [
        field.internalName,
        fieldLabel(messages, field.internalName, field.title),
      ]),
    );
    return {
      name: titleBy.get("FileLeafRef"),
      modified: titleBy.get("Modified"),
      size: titleBy.get("File_x0020_Size"),
    };
  }, [libraryFields, messages]);

  function openFolder(item: SharePointItem) {
    if (item.type !== "folder") return;
    setCrumbs((prev) => [...prev, { id: item.id, name: item.name }]);
  }

  return (
    <div className={mergeClasses(styles.root, className)}>
      <div className={styles.commandBar}>
        <Breadcrumb size="large">
          {crumbs.map((crumb, index) => (
            <Fragment key={crumb.id}>
              {index > 0 ? <BreadcrumbDivider /> : null}
              <BreadcrumbItem>
                <BreadcrumbButton
                  current={index === crumbs.length - 1}
                  onClick={() => setCrumbs((prev) => prev.slice(0, index + 1))}
                  icon={index === 0 ? <HomeRegular /> : undefined}
                >
                  {crumb.name}
                </BreadcrumbButton>
              </BreadcrumbItem>
            </Fragment>
          ))}
        </Breadcrumb>

        <div className={styles.commandActions}>
          {showLanguageSwitcher ? <LanguageSwitcher /> : null}
          {libraryFields.length > 0 ? (
            <ColumnPicker
              fields={selectableLibraryFields}
              visible={resolvedVisible}
              onVisibleChange={setVisibleExtra}
              label={messages.columns}
              messages={messages}
            />
          ) : null}
          <Button
            appearance="subtle"
            shape="circular"
            className={styles.commandIconButton}
            icon={<ArrowClockwiseRegular fontSize={20} />}
            onClick={() => void childrenQuery.refetch()}
            aria-label={messages.refresh}
            title={messages.refresh}
          />
        </div>
      </div>

      {childrenQuery.error ? (
        <ErrorBanner
          message={getErrorMessage(childrenQuery.error, messages.unknownError)}
          onRetry={() => void childrenQuery.refetch()}
          retryLabel={messages.retry}
        />
      ) : null}

      <div className={styles.listCard}>
        <div ref={rootRef} className={styles.listPane}>
          {childrenQuery.isPending ? <LibrarySkeleton /> : null}
          {!childrenQuery.isPending && items.length === 0 ? <EmptyState messages={messages} /> : null}
          {!childrenQuery.isPending && items.length > 0 ? (
            <FileList
              items={items}
              locale={locale}
              messages={messages}
              onOpenFolder={openFolder}
              extraColumns={extraColumns}
              fixedTitles={fixedTitles}
              sort={sort}
              onSort={onSort}
            />
          ) : null}
          {childrenQuery.hasNextPage ? (
            <div className={styles.loadMore}>
              <div ref={sentinelRef} className={styles.loadMoreSentinel} aria-hidden />
              <Button
                appearance="subtle"
                shape="circular"
                className={styles.loadMoreButton}
                disabled={childrenQuery.isFetchingNextPage}
                onClick={loadMore}
              >
                {messages.loadMore}
              </Button>
            </div>
          ) : null}
        </div>
      </div>
    </div>
  );
}
