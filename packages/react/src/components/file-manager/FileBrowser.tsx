import { Fragment, useCallback, useMemo, useState } from "react";
import {
  FIXED_LIBRARY_FIELD_NAMES,
  isSortableSearchField,
  type SharePointItem,
} from "@namphuongso/sharepoint-file-manager-core";
import {
  Breadcrumb,
  BreadcrumbButton,
  BreadcrumbDivider,
  BreadcrumbItem,
  Button,
  mergeClasses,
} from "@fluentui/react-components";
import {
  ArrowClockwiseRegular,
  DocumentSearchRegular,
  HomeRegular,
} from "@fluentui/react-icons";
import { useSharePoint } from "../../provider/context";
import { getErrorMessage } from "../../hooks/getErrorMessage";
import { useAccessibleItems } from "../../hooks/useAccessibleItems";
import { useColumnLayout } from "../../hooks/useColumnLayout";
import { useColumnSort } from "../../hooks/useColumnSort";
import { useFolderChildren } from "../../hooks/useFolderChildren";
import { useFolderViewCapabilities } from "../../hooks/useFolderViewCapabilities";
import { useLoadMoreOnScroll } from "../../hooks/useLoadMoreOnScroll";
import { useLibraryFields } from "../../hooks/useLibraryFields";
import { useVisibleExtraColumns } from "../../hooks/useVisibleExtraColumns";
import { fieldLabel } from "../../i18n/messages";
import type { BreadcrumbCrumb, FileBrowserProps, FileListColumn } from "../../types";
import { ColumnPicker } from "./ColumnPicker";
import { EmptyState } from "./EmptyState";
import { ErrorBanner } from "./ErrorBanner";
import { ForbiddenState } from "./ForbiddenState";
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

type BrowserView = "library" | "accessible";

/**
 * Duyệt thư viện (browse) + tab Search REST (item user được xem).
 * Không tạo client — phải nằm trong SharePointProvider.
 */
export function FileBrowser({ className, title, showLanguageSwitcher = true }: FileBrowserProps) {
  const styles = useFileManagerStyles();
  const { client, locale, messages } = useSharePoint();
  const rootId = client.config.rootItemId;
  const [view, setView] = useState<BrowserView>("library");
  const [crumbs, setCrumbs] = useState<BreadcrumbCrumb[]>([
    { id: rootId, name: title ?? messages.files },
  ]);
  const currentFolderId = crumbs[crumbs.length - 1]?.id ?? rootId;
  const columnScope = `${client.config.siteId}:${client.cacheScope}`;
  const librarySort = useColumnSort(columnScope);
  const accessibleSort = useColumnSort(`${columnScope}:accessible`);
  const isLibrary = view === "library";
  const { sort, onSort } = isLibrary ? librarySort : accessibleSort;

  const viewAccess = useFolderViewCapabilities(currentFolderId);
  const childrenQuery = useFolderChildren(currentFolderId, sort, {
    enabled: isLibrary && viewAccess.isReady && viewAccess.canView,
  });
  const fieldsQuery = useLibraryFields({
    enabled:
      view === "accessible" || (isLibrary && viewAccess.isReady && viewAccess.canView),
  });
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

  const visibleExtraNames = useMemo(
    () => extraNames.filter((name) => resolvedVisible.has(name)),
    [extraNames, resolvedVisible],
  );

  const accessibleQuery = useAccessibleItems(sort, {
    enabled: view === "accessible",
    fieldInternalNames: visibleExtraNames,
  });

  const libraryItems = useMemo(
    () => childrenQuery.data?.pages.flatMap((page) => page.items) ?? [],
    [childrenQuery.data],
  );
  const accessibleItems = useMemo(
    () => accessibleQuery.data?.pages.flatMap((page) => page.items) ?? [],
    [accessibleQuery.data],
  );
  const items = isLibrary ? libraryItems : accessibleItems;

  const fetchNextLibrary = childrenQuery.fetchNextPage;
  const fetchNextAccessible = accessibleQuery.fetchNextPage;
  const loadMore = useCallback(() => {
    if (isLibrary) void fetchNextLibrary();
    else void fetchNextAccessible();
  }, [isLibrary, fetchNextLibrary, fetchNextAccessible]);

  const hasNextPage = isLibrary ? childrenQuery.hasNextPage : accessibleQuery.hasNextPage;
  const isFetchingNextPage = isLibrary
    ? childrenQuery.isFetchingNextPage
    : accessibleQuery.isFetchingNextPage;

  const { rootRef, sentinelRef } = useLoadMoreOnScroll(
    Boolean(hasNextPage),
    isFetchingNextPage,
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
          kind: "extra" as const,
        })),
    [selectableLibraryFields, messages, resolvedVisible],
  );

  /** Nhãn 3 cột cố định + cột option (cả browse và Có quyền xem). */
  const defaultColumns = useMemo((): FileListColumn[] => {
    const titleBy = new Map(
      libraryFields.map((field) => [
        field.internalName,
        fieldLabel(messages, field.internalName, field.title),
      ]),
    );
    return [
      {
        internalName: "FileLeafRef",
        title: titleBy.get("FileLeafRef") ?? messages.name,
        typeAsString: "Text",
        kind: "name",
      },
      {
        internalName: "Modified",
        title: titleBy.get("Modified") ?? messages.modified,
        typeAsString: "DateTime",
        kind: "modified",
      },
      ...extraColumns,
      {
        internalName: "File_x0020_Size",
        title: titleBy.get("File_x0020_Size") ?? messages.size,
        typeAsString: "Number",
        kind: "size",
      },
    ];
  }, [libraryFields, messages, extraColumns]);

  const columnIds = useMemo(() => defaultColumns.map((col) => col.internalName), [defaultColumns]);
  const { order, widths, onReorder, onResize, onResizeEnd } = useColumnLayout(columnScope, columnIds);
  const columns = useMemo(() => {
    const byId = new Map(defaultColumns.map((col) => [col.internalName, col]));
    return order.flatMap((id) => {
      const col = byId.get(id);
      return col ? [col] : [];
    });
  }, [defaultColumns, order]);

  function openFolder(item: SharePointItem) {
    if (item.type !== "folder") return;
    if (!isLibrary) {
      // Folder từ Search → nhảy sang browse đúng UniqueId.
      setView("library");
      setCrumbs([
        { id: rootId, name: title ?? messages.files },
        { id: item.id, name: item.name },
      ]);
      return;
    }
    setCrumbs((prev) => [...prev, { id: item.id, name: item.name }]);
  }

  function refreshAll() {
    if (isLibrary) {
      void viewAccess.refetch();
      void childrenQuery.refetch();
      return;
    }
    void accessibleQuery.refetch();
  }

  const listLoading = isLibrary
    ? viewAccess.isLoading || (viewAccess.canView && childrenQuery.isPending)
    : accessibleQuery.isPending;

  const showLibraryForbidden = isLibrary && !listLoading && viewAccess.viewDenied;
  const showEmpty =
    !listLoading &&
    !showLibraryForbidden &&
    items.length === 0 &&
    (isLibrary ? viewAccess.canView : !accessibleQuery.isError);
  const showList =
    !listLoading &&
    !showLibraryForbidden &&
    items.length > 0 &&
    (isLibrary ? viewAccess.canView : true);

  return (
    <div className={mergeClasses(styles.root, className)}>
      <div className={styles.commandBar}>
        <div className={styles.commandBarStart}>
          <div className={styles.viewSwitch} role="tablist" aria-label={messages.files}>
            <Button
              appearance="subtle"
              shape="circular"
              className={mergeClasses(
                styles.commandIconButton,
                isLibrary && styles.commandIconButtonActive,
              )}
              icon={<HomeRegular fontSize={20} />}
              onClick={() => setView("library")}
              aria-label={messages.tabLibrary}
              title={messages.tabLibrary}
              aria-selected={isLibrary}
              role="tab"
            />
            <Button
              appearance="subtle"
              shape="circular"
              className={mergeClasses(
                styles.commandIconButton,
                !isLibrary && styles.commandIconButtonActive,
              )}
              icon={<DocumentSearchRegular fontSize={20} />}
              onClick={() => setView("accessible")}
              aria-label={messages.tabAccessible}
              title={messages.tabAccessible}
              aria-selected={!isLibrary}
              role="tab"
            />
          </div>

          {isLibrary ? (
            <Breadcrumb size="large">
              {crumbs.map((crumb, index) => (
                <Fragment key={crumb.id}>
                  {index > 0 ? <BreadcrumbDivider /> : null}
                  <BreadcrumbItem>
                    <BreadcrumbButton
                      current={index === crumbs.length - 1}
                      onClick={() => setCrumbs((prev) => prev.slice(0, index + 1))}
                    >
                      {crumb.name}
                    </BreadcrumbButton>
                  </BreadcrumbItem>
                </Fragment>
              ))}
            </Breadcrumb>
          ) : (
            <Breadcrumb size="large">
              <BreadcrumbItem>
                <BreadcrumbButton current>{messages.tabAccessible}</BreadcrumbButton>
              </BreadcrumbItem>
            </Breadcrumb>
          )}
        </div>

        <div className={styles.commandActions}>
          {showLanguageSwitcher ? <LanguageSwitcher /> : null}
          {libraryFields.length > 0 &&
          (isLibrary ? viewAccess.canView : !accessibleQuery.isError) ? (
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
            onClick={refreshAll}
            aria-label={messages.refresh}
            title={messages.refresh}
          />
        </div>
      </div>

      {isLibrary && viewAccess.isError ? (
        <ErrorBanner
          message={getErrorMessage(viewAccess.error, messages.unknownError)}
          onRetry={() => void viewAccess.refetch()}
          retryLabel={messages.retry}
        />
      ) : null}

      {isLibrary && childrenQuery.error && viewAccess.canView ? (
        <ErrorBanner
          message={getErrorMessage(childrenQuery.error, messages.unknownError)}
          onRetry={() => void childrenQuery.refetch()}
          retryLabel={messages.retry}
        />
      ) : null}

      {!isLibrary && accessibleQuery.error ? (
        <ErrorBanner
          message={getErrorMessage(accessibleQuery.error, messages.unknownError)}
          onRetry={() => void accessibleQuery.refetch()}
          retryLabel={messages.retry}
        />
      ) : null}

      <div className={styles.listCard}>
        <div ref={rootRef} className={styles.listPane}>
          {listLoading ? <LibrarySkeleton /> : null}
          {showLibraryForbidden ? <ForbiddenState messages={messages} /> : null}
          {showEmpty ? (
            <EmptyState
              messages={messages}
              title={isLibrary ? undefined : messages.searchEmpty}
              hint={isLibrary ? undefined : messages.searchEmptyHint}
            />
          ) : null}
          {showList ? (
            <FileList
              items={items}
              locale={locale}
              messages={messages}
              onOpenFolder={openFolder}
              columns={columns}
              columnWidths={widths}
              onColumnResize={onResize}
              onColumnResizeEnd={onResizeEnd}
              onColumnReorder={onReorder}
              sort={sort}
              onSort={onSort}
              isSortable={isLibrary ? undefined : isSortableSearchField}
            />
          ) : null}
          {(isLibrary ? viewAccess.canView : !accessibleQuery.isError) && hasNextPage ? (
            <div className={styles.loadMore}>
              <div ref={sentinelRef} className={styles.loadMoreSentinel} aria-hidden />
              <Button
                appearance="subtle"
                shape="circular"
                className={styles.loadMoreButton}
                disabled={isFetchingNextPage}
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
