import { Fragment, useCallback, useEffect, useMemo, useRef, useState, type ChangeEvent, type MouseEvent } from "react";
import {
  FIXED_LIBRARY_FIELD_NAMES,
  isSortableSearchField,
  NEW_DOCUMENT_KINDS,
  type NewDocumentKind,
  type SharePointItem,
} from "@namphuongso/sharepoint-file-manager-core";
import {
  Breadcrumb,
  BreadcrumbButton,
  BreadcrumbDivider,
  BreadcrumbItem,
  Button,
  MessageBar,
  MessageBarBody,
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
import { useCreateDocument } from "../../hooks/useCreateDocument";
import { useCreateFolder } from "../../hooks/useCreateFolder";
import { useDeleteItem, isDeleteDenied } from "../../hooks/useDeleteItem";
import { useFileBrowserNavigation } from "../../hooks/useFileBrowserNavigation";
import { useFolderChildren } from "../../hooks/useFolderChildren";
import { useFolderViewCapabilities } from "../../hooks/useFolderViewCapabilities";
import { useLoadMoreOnScroll } from "../../hooks/useLoadMoreOnScroll";
import { useLibraryFields } from "../../hooks/useLibraryFields";
import { useOpenItem } from "../../hooks/useOpenItem";
import { useDownloadItem } from "../../hooks/useDownloadItem";
import { useUploadFile } from "../../hooks/useUploadFile";
import { useUploadFolder } from "../../hooks/useUploadFolder";
import { useVisibleExtraColumns } from "../../hooks/useVisibleExtraColumns";
import { fieldLabel } from "../../i18n/messages";
import type { FileBrowserProps, FileListColumn } from "../../types";
import { ColumnPicker } from "./ColumnPicker";
import { CreateDocumentDialog } from "./CreateDocumentDialog";
import { CreateFolderDialog } from "./CreateFolderDialog";
import { DeleteItemDialog } from "./DeleteItemDialog";
import { EmptyState } from "./EmptyState";
import { ErrorBanner } from "./ErrorBanner";
import { ForbiddenState } from "./ForbiddenState";
import { FileList } from "./FileView";
import { ItemContextMenu } from "./ItemContextMenu";
import { LanguageSwitcher } from "./LanguageSwitcher";
import { LibrarySkeleton } from "./LibrarySkeleton";
import { ListContextMenu } from "./ListContextMenu";
import type { NewItemAction } from "./newItemActions";
import { NewItemToolbarMenu } from "./NewItemToolbarMenu";
import { useFileManagerStyles } from "./useFileManagerStyles";

/** ItemChildCount: hiện Folder.ItemCount. FolderChildCount không có số tách — không đưa picker. */
const ITEM_CHILD_COUNT = "ItemChildCount";
const HIDE_FROM_PICKER = new Set([
  "FolderChildCount",
  "ComplianceAssetId",
  "DocIcon",
  "FileSize",
]);

/** Mở OS picker đồng bộ trong user gesture (MenuItem onClick). */
function openFilePicker(input: HTMLInputElement | null) {
  if (!input) return;
  try {
    if (typeof input.showPicker === "function") {
      void input.showPicker();
      return;
    }
  } catch {
    // showPicker bị chặn / không hỗ trợ → click().
  }
  input.click();
}

/**
 * Duyệt thư viện (browse) + tab Search REST (item user được xem).
 * Không tạo client — phải nằm trong SharePointProvider.
 */
export function FileBrowser({ className, title, showLanguageSwitcher = true }: FileBrowserProps) {
  const styles = useFileManagerStyles();
  const { client, locale, messages } = useSharePoint();
  const rootId = client.config.rootItemId;
  const rootName = title ?? messages.files;
  const { view, setView, crumbs, setCrumbs, navigate, locationReady } =
    useFileBrowserNavigation(rootId, rootName);
  const currentFolderId = crumbs[crumbs.length - 1]?.id ?? rootId;
  const columnScope = `${client.config.siteId}:${client.cacheScope}`;
  const librarySort = useColumnSort(columnScope);
  const accessibleSort = useColumnSort(`${columnScope}:accessible`);
  const isLibrary = view === "library";
  const { sort, onSort } = isLibrary ? librarySort : accessibleSort;

  const viewAccess = useFolderViewCapabilities(locationReady ? currentFolderId : undefined);
  const childrenQuery = useFolderChildren(currentFolderId, sort, {
    enabled: isLibrary && locationReady && viewAccess.isReady && viewAccess.canView,
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

  const [fileActionState, setFileActionState] = useState<{
    kind: "denied" | "error";
    action: "open" | "download" | "delete";
    message?: string;
  }>();
  const [folderDialogOpen, setFolderDialogOpen] = useState(false);
  const [folderDialogError, setFolderDialogError] = useState<string>();
  const [documentKind, setDocumentKind] = useState<NewDocumentKind | null>(null);
  const [documentDialogError, setDocumentDialogError] = useState<string>();
  const [deleteTarget, setDeleteTarget] = useState<SharePointItem | null>(null);
  const [contextMenu, setContextMenu] = useState<{ x: number; y: number } | null>(null);
  const [itemMenu, setItemMenu] = useState<{
    item: SharePointItem;
    x: number;
    y: number;
  } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const el = folderInputRef.current;
    if (!el) return;
    el.setAttribute("webkitdirectory", "");
    el.setAttribute("directory", "");
  }, []);

  const createFolder = useCreateFolder(currentFolderId);
  const createDocument = useCreateDocument(currentFolderId);
  const uploadFile = useUploadFile(currentFolderId);
  const uploadFolder = useUploadFolder(currentFolderId);
  const deleteItem = useDeleteItem(currentFolderId);
  const canWrite = isLibrary && viewAccess.canView && viewAccess.canAdd;
  const writeBusy =
    createFolder.isPending ||
    createDocument.isPending ||
    uploadFile.isPending ||
    uploadFolder.isPending;
  const { openItem, openingId } = useOpenItem();
  const { downloadItem, downloadingId } = useDownloadItem();
  const itemActionBusy = Boolean(openingId || downloadingId || deleteItem.isPending);

  const onOpenFile = useCallback(
    (item: SharePointItem) => {
      setFileActionState(undefined);
      void openItem(item).then((result) => {
        if (result.status === "opened") {
          // 2 pha: mở trước (được coi như user gesture), lỗi thì hiển thị banner.
          window.open(result.url, "_blank", "noopener,noreferrer");
          return;
        }
        if (result.status === "denied") {
          setFileActionState({ kind: "denied", action: "open" });
          return;
        }
        setFileActionState({ kind: "error", action: "open", message: result.message });
      });
    },
    [openItem],
  );

  const onDownloadFile = useCallback(
    (item: SharePointItem) => {
      setFileActionState(undefined);
      void downloadItem(item).then((result) => {
        // Toast đã báo success/error; banner chỉ khi thiếu OpenItems.
        if (result.status === "denied") {
          setFileActionState({ kind: "denied", action: "download" });
        }
      });
    },
    [downloadItem],
  );

  const onDeleteRequest = useCallback((item: SharePointItem) => {
    setFileActionState(undefined);
    setDeleteTarget(item);
  }, []);

  const handleConfirmDelete = useCallback(async () => {
    if (!deleteTarget) return;
    try {
      await deleteItem.mutateAsync(deleteTarget);
      setDeleteTarget(null);
    } catch (error) {
      if (isDeleteDenied(error)) {
        setDeleteTarget(null);
        setFileActionState({ kind: "denied", action: "delete" });
      }
      // Lỗi khác: toast đã báo; giữ dialog để user hủy hoặc thử lại.
    }
  }, [deleteItem, deleteTarget]);

  const onItemContextMenu = useCallback((item: SharePointItem, position: { x: number; y: number }) => {
    setContextMenu(null);
    setItemMenu({ item, ...position });
  }, []);

  const documentDefaultName = useMemo(() => {
    if (!documentKind) return "";
    const info = NEW_DOCUMENT_KINDS.find((k) => k.kind === documentKind);
    return info?.defaultBaseName ?? "Document";
  }, [documentKind]);

  function handleNewAction(action: NewItemAction) {
    if (!canWrite || writeBusy) return;
    if (action.type === "folder") {
      setFolderDialogError(undefined);
      setFolderDialogOpen(true);
      return;
    }
    if (action.type === "uploadFiles" || action.type === "uploadFolder") return;
    setDocumentDialogError(undefined);
    setDocumentKind(action.kind);
  }

  function handlePickFiles() {
    if (!canWrite || writeBusy) return;
    openFilePicker(fileInputRef.current);
  }

  function handlePickFolder() {
    if (!canWrite || writeBusy) return;
    openFilePicker(folderInputRef.current);
  }

  function handleListContextMenu(event: MouseEvent) {
    if (!isLibrary || !viewAccess.canView) return;
    // Chỉ nền khung / empty — chuột phải trên dòng đã stopPropagation trong FileRow.
    const target = event.target as HTMLElement;
    if (target.closest("[data-file-row]")) return;
    event.preventDefault();
    setItemMenu(null);
    setContextMenu({ x: event.clientX, y: event.clientY });
  }

  async function handleCreateFolder(name: string) {
    setFolderDialogError(undefined);
    try {
      await createFolder.mutateAsync(name);
      setFolderDialogOpen(false);
    } catch (error) {
      // Lỗi validate (tên trống / trùng) hiện trong dialog; lỗi POST đã có toast từ hook.
      setFolderDialogError(
        error instanceof Error && error.message ? error.message : messages.createFolderError,
      );
    }
  }

  async function handleCreateDocument(values: { name: string }) {
    if (!documentKind) return;
    setDocumentDialogError(undefined);
    try {
      await createDocument.mutateAsync({
        kind: documentKind,
        name: values.name,
      });
      setDocumentKind(null);
    } catch (error) {
      setDocumentDialogError(
        error instanceof Error && error.message ? error.message : messages.createDocumentError,
      );
    }
  }

  async function handleUploadChange(event: ChangeEvent<HTMLInputElement>) {
    // FileList là live — copy trước khi reset value (reset làm length = 0).
    const files = event.target.files ? [...event.target.files] : [];
    event.target.value = "";
    if (files.length === 0) return;
    for (const file of files) {
      try {
        await uploadFile.mutateAsync({ file });
      } catch {
        // Lỗi đã có toast từ hook — không nuốt nhưng cũng không dừng cả batch.
      }
    }
  }

  async function handleFolderUploadChange(event: ChangeEvent<HTMLInputElement>) {
    const files = event.target.files ? [...event.target.files] : [];
    event.target.value = "";
    if (files.length === 0) return;
    try {
      await uploadFolder.mutateAsync(files);
    } catch {
      // Lỗi đã có toast từ hook.
    }
  }

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
      // Folder từ Search → nhảy sang browse đúng UniqueId + URL.
      navigate(
        [
          { id: rootId, name: rootName },
          { id: item.id, name: item.name },
        ],
        "library",
      );
      return;
    }
    setCrumbs((prev) => [...prev, { id: item.id, name: item.name }]);
  }

  function refreshAll() {
    if (isRefreshing) return;
    if (isLibrary) {
      void viewAccess.refetch();
      void childrenQuery.refetch();
      return;
    }
    void accessibleQuery.refetch();
  }

  /** Nút refresh đang quay khi bất kỳ query list nào fetching (cả lần đầu lẫn refetch). */
  const isRefreshing = isLibrary
    ? viewAccess.isFetching || childrenQuery.isFetching
    : accessibleQuery.isFetching;

  const listLoading =
    !locationReady ||
    (isLibrary
      ? viewAccess.isLoading || (viewAccess.canView && childrenQuery.isPending)
      : accessibleQuery.isPending);

  /** Skeleton cũng hiện khi user bấm refresh (đã có data, đang refetch).
   *  Lần đầu (chưa có data) đã được `listLoading` cover. */
  const isRefetchingList = isLibrary
    ? childrenQuery.isFetching && !childrenQuery.isPending
    : accessibleQuery.isFetching && !accessibleQuery.isPending;
  /** Loading tổng: lần đầu HOẶC đang refetch list — cả hai đều ưu tiên skeleton. */
  const showSkeleton = listLoading || isRefetchingList;

  const showLibraryForbidden = isLibrary && !showSkeleton && viewAccess.viewDenied;
  const showEmpty =
    !showSkeleton &&
    !showLibraryForbidden &&
    items.length === 0 &&
    (isLibrary ? viewAccess.canView : !accessibleQuery.isError);
  const showList =
    !showSkeleton &&
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
              onClick={() => navigate([{ id: rootId, name: rootName }], "library")}
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

          {isLibrary && viewAccess.canView && canWrite ? (
            <div className={styles.viewSwitch} role="group" aria-label={messages.newItem}>
              <NewItemToolbarMenu
                messages={messages}
                disabled={writeBusy}
                onAction={handleNewAction}
                onPickFiles={handlePickFiles}
                onPickFolder={handlePickFolder}
              />
            </div>
          ) : null}

          {/* Input luôn mount ngoài Menu — tránh unmount khi popover đóng. */}
          <input
            ref={fileInputRef}
            type="file"
            multiple
            className={styles.visuallyHiddenInput}
            tabIndex={-1}
            onChange={(e) => void handleUploadChange(e)}
          />
          <input
            ref={folderInputRef}
            type="file"
            className={styles.visuallyHiddenInput}
            tabIndex={-1}
            onChange={(e) => void handleFolderUploadChange(e)}
          />

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
            className={mergeClasses(
              styles.commandIconButton,
              isRefreshing && styles.commandIconButtonSpinning,
            )}
            icon={<ArrowClockwiseRegular fontSize={20} />}
            onClick={refreshAll}
            disabled={isRefreshing}
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

      {fileActionState ? (
        fileActionState.kind === "denied" ? (
          <MessageBar intent="warning" className={styles.errorBanner}>
            <MessageBarBody>
              {fileActionState.action === "delete"
                ? messages.noDeletePermission
                : `${messages.noOpenPermission} — ${messages.noOpenPermissionHint}`}
            </MessageBarBody>
          </MessageBar>
        ) : (
          <ErrorBanner
            message={`${
              fileActionState.action === "download"
                ? messages.downloadError
                : fileActionState.action === "delete"
                  ? messages.deleteError
                  : messages.openFileError
            }: ${fileActionState.message ?? messages.unknownError}`}
            retryLabel={messages.refresh}
          />
        )
      ) : null}

      <CreateFolderDialog
        open={folderDialogOpen}
        onOpenChange={setFolderDialogOpen}
        onSubmit={handleCreateFolder}
        messages={messages}
        isPending={createFolder.isPending}
        errorMessage={folderDialogError}
      />

      <CreateDocumentDialog
        open={documentKind !== null}
        kind={documentKind}
        onOpenChange={(open) => {
          if (!open) setDocumentKind(null);
        }}
        onSubmit={handleCreateDocument}
        messages={messages}
        defaultName={documentDefaultName}
        isPending={createDocument.isPending}
        errorMessage={documentDialogError}
      />

      <DeleteItemDialog
        open={deleteTarget !== null}
        item={deleteTarget}
        onOpenChange={(open) => {
          if (!open) setDeleteTarget(null);
        }}
        onConfirm={handleConfirmDelete}
        messages={messages}
        isPending={deleteItem.isPending}
      />

      <ListContextMenu
        open={contextMenu !== null}
        anchor={contextMenu}
        onOpenChange={(open) => {
          if (!open) setContextMenu(null);
        }}
        messages={messages}
        canAdd={canWrite}
        writeBusy={writeBusy}
        onAction={handleNewAction}
        onPickFiles={handlePickFiles}
        onPickFolder={handlePickFolder}
        onRefresh={refreshAll}
      />

      <ItemContextMenu
        open={itemMenu !== null}
        anchor={itemMenu ? { x: itemMenu.x, y: itemMenu.y } : null}
        item={itemMenu?.item ?? null}
        onOpenChange={(open) => {
          if (!open) setItemMenu(null);
        }}
        messages={messages}
        busy={itemActionBusy}
        onOpen={(item) => {
          if (item.type === "folder") openFolder(item);
          else onOpenFile(item);
        }}
        onDownload={onDownloadFile}
        onDelete={onDeleteRequest}
      />

      <div className={styles.listCard}>
        <div
          ref={rootRef}
          className={styles.listPane}
          onContextMenu={handleListContextMenu}
        >
          {showSkeleton ? <LibrarySkeleton /> : null}
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
              onOpenFile={onOpenFile}
              onDownloadFile={onDownloadFile}
              onDeleteFile={onDeleteRequest}
              onItemContextMenu={onItemContextMenu}
              itemActionBusy={itemActionBusy}
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
