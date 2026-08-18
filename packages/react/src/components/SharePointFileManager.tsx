import { useMemo, useRef, useState, useEffect } from "react";
import type {
  ConflictBehavior,
  CopyOperationProgress,
  NotifyPayload,
  OfficeFileKind,
  PreviewInfo,
  SharePointConfig,
  SharePointItem,
  SortDirection,
  SortField,
  SearchScope,
  SearchFilters,
} from "@namphuongso/sharepoint-file-manager-core";
import { decodeLibraryPage } from "@namphuongso/sharepoint-file-manager-core";
import { SharePointProvider } from "../provider/SharePointProvider";
import { useSharePoint } from "../provider/context";
import {
  getErrorMessage,
  useCheckout,
  useCopyItem,
  useCreateFolder,
  useCreateLink,
  useCreateOfficeFile,
  useDeleteItem,
  useDownloadFile,
  useDownloadVersion,
  useFolderChildren,
  useFolderChildrenInfinite,
  useInvite,
  useItem,
  useItemActivities,
  useListColumns,
  useListItemFields,
  useMoveItem,
  usePermissions,
  useRemovePermission,
  useRenameItem,
  useRestoreVersion,
  useSearchItems,
  useUpdatePermission,
  useUpdateListItemFields,
  useBulkUpdateListItemFields,
  useUploadFile,
  useVersions,
} from "../hooks/hooks";
import { useKeyboardShortcuts } from "../hooks/useKeyboardShortcuts";
import type { FileAction } from "./ContextMenu";
import { ContextMenu } from "./ContextMenu";
import { FileGrid, FileList } from "./FileView";
import { getFileKind } from "./FileTypeIcon";
import { CommandBar, type FileTypeFilter } from "./CommandBar";
import { DetailsPane } from "./DetailsPane";
import { EmptyState } from "./EmptyState";
import { LibrarySkeleton } from "./LibrarySkeleton";
import { CopyMoveDialog, CreateFolderDialog, DeleteDialog, RenameDialog, CheckinDialog } from "./dialogs/BasicDialogs";
import { ShareDialog, UploadDialog } from "./dialogs/ShareUploadDialogs";
import type { SelectionAction } from "./SelectionToolbar";
import { ManageAccessDialog, PreviewDialog, PropertiesDialog, VersionHistoryDialog } from "./dialogs/AccessDialogs";
import { Button, ErrorBanner } from "./ui";
import { FilterPanel } from "./FilterPanel";
import { ActivityDialog } from "./ActivityDialog";
import { BulkMetadataDialog } from "./dialogs/BulkMetadataDialog";
import { useNotify } from "../hooks/useNotify";
import { ColumnChooser } from "./ColumnChooser";
import {
  Breadcrumb,
  BreadcrumbButton,
  BreadcrumbItem,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Text,
  tokens,
} from "@fluentui/react-components";
import { FolderRegular, MoreHorizontalRegular } from "@fluentui/react-icons";
import {
  loadColumnSettings,
  saveColumnSettings,
  defaultMetadataColumnNames,
  type ColumnVisibilitySettings,
} from "../utils/column-settings";
import type { Messages } from "../i18n/messages";

const DRAG_MIME = "application/x-spm-items";

export interface SharePointFileManagerProps {
  config: SharePointConfig;
  locale?: string;
  view?: "list" | "compact" | "grid";
  className?: string;
  title?: string;
  messages?: Partial<Messages>;
  onNotify?: (payload: NotifyPayload) => void;
  theme?: "light" | "dark" | "system";
  density?: "compact" | "comfortable";
  embedded?: boolean;
  showHeader?: boolean;
  showNavigation?: boolean;
  showBreadcrumb?: boolean;
}

export function SharePointFileManager(props: SharePointFileManagerProps) {
  return (
    <SharePointProvider
      config={props.config}
      locale={props.locale}
      messages={props.messages}
      onNotify={props.onNotify}
      theme={props.theme}
      density={props.density}
      embedded={props.embedded}
    >
      <FileManagerShell initialView={props.view ?? "list"} className={props.className} title={props.title} density={props.density ?? "comfortable"} />
    </SharePointProvider>
  );
}

interface Crumb {
  id: string;
  name: string;
}

function FileManagerShell({
  initialView,
  className,
  title,
  density,
}: {
  initialView: "list" | "compact" | "grid";
  className?: string;
  title?: string;
  density: "compact" | "comfortable";
}) {
  const { client, locale, messages } = useSharePoint();
  const features = client.config.features;
  const notify = useNotify();
  const rootId = client.config.rootItemId;
  const quickAccessScope = `${client.config.siteId}:${rootId}`;
  const [crumbs, setCrumbs] = useState<Crumb[]>([{ id: rootId, name: messages.files }]);
  const currentFolderId = crumbs[crumbs.length - 1]?.id ?? rootId;
  const visibleCrumbs: Crumb[] = crumbs.length > 4 ? [crumbs[0]!, ...crumbs.slice(-2)] : crumbs;
  const hiddenCrumbs = crumbs.length > 4 ? crumbs.slice(1, -2) : [];
  const viewSettingsKey = `spm-view:${quickAccessScope}`;
  const [view, setView] = useState<"list" | "compact" | "grid">(() => {
    if (typeof localStorage === "undefined") return initialView;
    const stored = localStorage.getItem(viewSettingsKey);
    return stored === "list" || stored === "compact" || stored === "grid" ? stored : initialView;
  });
  const [query, setQuery] = useState("");
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [menu, setMenu] = useState<{ item: SharePointItem; x: number; y: number } | null>(null);
  const [dialog, setDialog] = useState<
    | { type: "createFolder" }
    | { type: "upload" }
    | { type: "rename"; item: SharePointItem }
    | { type: "delete"; items: SharePointItem[] }
    | { type: "copy"; items: SharePointItem[] }
    | { type: "move"; items: SharePointItem[] }
    | { type: "share"; item: SharePointItem }
    | { type: "access"; item: SharePointItem }
    | { type: "preview"; item: SharePointItem }
    | { type: "versions"; item: SharePointItem }
    | { type: "properties"; item: SharePointItem }
    | { type: "checkin"; item: SharePointItem }
    | { type: "activity"; item: SharePointItem }
    | { type: "bulkMetadata"; items: SharePointItem[] }
    | null
  >(null);
  const [searchScope, setSearchScope] = useState<SearchScope>("folder");
  const [searchFilters, setSearchFilters] = useState<SearchFilters>({});
  const [appliedFilters, setAppliedFilters] = useState<SearchFilters>({});
  const [showFilters, setShowFilters] = useState(false);
  const [showColumnChooser, setShowColumnChooser] = useState(false);
  const [detailsOpen, setDetailsOpen] = useState(false);
  const [typeFilter, setTypeFilter] = useState<FileTypeFilter | undefined>();
  const [inlineRenameId, setInlineRenameId] = useState<string | undefined>();
  const columnSettingsScope = `${client.config.siteId}:${client.config.listId ?? rootId}`;
  const [columnSettings, setColumnSettings] = useState<ColumnVisibilitySettings>(() =>
    loadColumnSettings(columnSettingsScope),
  );
  const [sortDirection, setSortDirection] = useState<SortDirection>("asc");
  const lastSelectedIndexRef = useRef<number>(-1);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [preview, setPreview] = useState<PreviewInfo | undefined>();
  const [actionError, setActionError] = useState<string | undefined>();
  const [grantSuccess, setGrantSuccess] = useState(false);
  const [restoreSuccess, setRestoreSuccess] = useState(false);
  const [searchExtraItems, setSearchExtraItems] = useState<SharePointItem[]>([]);
  const [searchNextLink, setSearchNextLink] = useState<string | undefined>();
  const [loadingMoreSearch, setLoadingMoreSearch] = useState(false);
  const [copyProgress, setCopyProgress] = useState<number | undefined>();
  const [breadcrumbDropTargetId, setBreadcrumbDropTargetId] = useState<string | null>(null);
  const [isDragOver, setIsDragOver] = useState(false);
  const dropRef = useRef<HTMLDivElement>(null);
  const folderUploadRef = useRef<HTMLInputElement>(null);
  const uploadAbortRef = useRef<AbortController | null>(null);

  useEffect(() => {
    if (typeof localStorage !== "undefined") localStorage.setItem(viewSettingsKey, view);
  }, [view, viewSettingsKey]);

  const isSearching = query.trim().length >= 2;
  const useInfiniteListing = client.config.features.infiniteScroll && !isSearching;
  const childrenQuery = useFolderChildren(
    useInfiniteListing ? undefined : currentFolderId,
    features.metadata,
  );
  const infiniteChildrenQuery = useFolderChildrenInfinite(
    useInfiniteListing ? currentFolderId : undefined,
    client.config.features.metadata,
  );
  const [sortField, setSortField] = useState<SortField>("name");
  const listColumnsQuery = useListColumns();
  const allMetadataColumns = listColumnsQuery.data ?? [];
  const visibleMetadataColumns = useMemo(() => {
    const names = columnSettings.metadataColumnNames;
    const effectiveNames =
      names.length > 0 ? names : defaultMetadataColumnNames(allMetadataColumns.map((column) => column.name));
    return allMetadataColumns.filter((column) => effectiveNames.includes(column.name));
  }, [allMetadataColumns, columnSettings.metadataColumnNames]);

  useEffect(() => {
    setColumnSettings(loadColumnSettings(columnSettingsScope));
  }, [columnSettingsScope]);

  useEffect(() => {
    setInlineRenameId(undefined);
  }, [currentFolderId, selectedIds.length]);

  function updateColumnSettings(next: ColumnVisibilitySettings) {
    setColumnSettings(next);
    saveColumnSettings(columnSettingsScope, next);
  }

  async function handleInlineRename(item: SharePointItem, name: string) {
    setInlineRenameId(undefined);
    setActionError(undefined);
    try {
      await renameItem.mutateAsync({ itemId: item.id, name });
      reportSuccess(messages.operationSuccess);
    } catch (caught) {
      reportError(caught);
    }
  }
  const searchQuery = useSearchItems(query, {
    folderId: currentFolderId,
    scope: searchScope,
    filters: appliedFilters,
  });

  useEffect(() => {
    setSearchExtraItems([]);
    setSearchNextLink(undefined);
  }, [query, currentFolderId, searchScope, appliedFilters]);

  useEffect(() => {
    if (searchQuery.data?.nextLink) {
      setSearchNextLink(searchQuery.data.nextLink);
    } else if (!isSearching) {
      setSearchNextLink(undefined);
    }
  }, [searchQuery.data, isSearching]);

  const folderItems = useMemo(() => {
    if (useInfiniteListing) {
      return infiniteChildrenQuery.data?.pages.flatMap((page) => page.items) ?? [];
    }
    return childrenQuery.data?.items ?? [];
  }, [useInfiniteListing, infiniteChildrenQuery.data, childrenQuery.data]);

  const items = useMemo(() => {
    if (isSearching) {
      const base = searchQuery.data?.items ?? [];
      return [...base, ...searchExtraItems];
    }
    return folderItems;
  }, [folderItems, isSearching, searchExtraItems, searchQuery.data]);

  const sortedItems = useMemo(() => {
    const dir = sortDirection === "asc" ? 1 : -1;
    return [...items].sort((a, b) => {
      if (a.type !== b.type) return a.type === "folder" ? -1 : 1;
      if (sortField === "modified") {
        return (a.lastModifiedDateTime ?? "").localeCompare(b.lastModifiedDateTime ?? "") * dir;
      }
      if (sortField === "created") {
        return (a.createdDateTime ?? "").localeCompare(b.createdDateTime ?? "") * dir;
      }
      if (sortField === "size") return ((a.size ?? 0) - (b.size ?? 0)) * dir;
      return a.name.localeCompare(b.name, undefined, { sensitivity: "base" }) * dir;
    });
  }, [items, sortDirection, sortField]);

  const visibleItems = useMemo(() => {
    if (!typeFilter) return sortedItems;
    return sortedItems.filter((item) => getFileKind(item) === typeFilter);
  }, [sortedItems, typeFilter]);

  const selectedItems = visibleItems.filter((item) => selectedIds.includes(item.id));
  const selectedItem = selectedItems[0] ?? menu?.item;
  const activeItem =
    dialog && "item" in dialog ? dialog.item : selectedItem;

  const createFolder = useCreateFolder(currentFolderId);
  const renameItem = useRenameItem(currentFolderId);
  const deleteItem = useDeleteItem(currentFolderId);
  const copyItem = useCopyItem(currentFolderId);
  const moveItem = useMoveItem(currentFolderId);
  const uploadFile = useUploadFile(currentFolderId);
  const downloadFile = useDownloadFile();
  const downloadVersion = useDownloadVersion();
  const checkoutMutation = useCheckout(currentFolderId);
  const createOfficeFile = useCreateOfficeFile(currentFolderId);
  const invite = useInvite(activeItem?.id ?? "");
  const accessInvite = useInvite(dialog?.type === "access" ? dialog.item.id : "");
  const createLink = useCreateLink(activeItem?.id ?? "");
  const removePermission = useRemovePermission(activeItem?.id ?? "");
  const updatePermission = useUpdatePermission(activeItem?.id ?? "");
  const restoreVersion = useRestoreVersion(currentFolderId);
  const permissionsQuery = usePermissions(dialog?.type === "access" ? dialog.item.id : undefined);
  const versionsQuery = useVersions(dialog?.type === "versions" ? dialog.item.id : undefined, dialog?.type === "versions");
  const propertiesQuery = useItem(dialog?.type === "properties" ? dialog.item.id : undefined, true);
  const listItemFieldsQuery = useListItemFields(
    dialog?.type === "properties" ? dialog.item.id : undefined,
    dialog?.type === "properties",
  );
  const updateListItemFields = useUpdateListItemFields(
    dialog?.type === "properties" ? dialog.item.id : "",
  );
  const bulkUpdateListItemFields = useBulkUpdateListItemFields();
  const activitiesQuery = useItemActivities(
    dialog?.type === "activity" ? dialog.item.id : undefined,
    dialog?.type === "activity",
  );

  function emitNotify(payload: NotifyPayload) {
    notify(payload);
  }

  function reportError(caught: unknown) {
    const message = getErrorMessage(caught, messages);
    setActionError(message);
    emitNotify({ type: "error", message });
    return message;
  }

  function reportSuccess(message: string) {
    emitNotify({ type: "success", message });
  }

  const loading = isSearching
    ? searchQuery.isLoading
    : useInfiniteListing
      ? infiniteChildrenQuery.isLoading
      : childrenQuery.isLoading;
  const error = isSearching
    ? searchQuery.error
    : useInfiniteListing
      ? infiniteChildrenQuery.error
      : childrenQuery.error;
  const folderNextLink = useInfiniteListing
    ? infiniteChildrenQuery.hasNextPage
      ? "more"
      : undefined
    : undefined;

  function refreshListing() {
    setActionError(undefined);
    if (isSearching) {
      setSearchExtraItems([]);
      void searchQuery.refetch();
      return;
    }
    if (useInfiniteListing) {
      void infiniteChildrenQuery.refetch();
      return;
    }
    void childrenQuery.refetch();
  }

  async function loadMoreSearch() {
    if (!searchNextLink || loadingMoreSearch) return;
    setLoadingMoreSearch(true);
    try {
      const result =
        searchScope === "library"
          ? await client.search.search({
              query: query.trim(),
              scope: "library",
              folderId: currentFolderId,
              filters: appliedFilters,
              from: decodeLibraryPage(searchNextLink),
            })
          : await client.search.search({
              query: query.trim(),
              folderId: currentFolderId,
              nextLink: searchNextLink,
            });
      setSearchExtraItems((current) => [...current, ...result.items]);
      setSearchNextLink(result.nextLink);
    } catch (caught) {
      setActionError(getErrorMessage(caught, messages));
    } finally {
      setLoadingMoreSearch(false);
    }
  }

  function selectAllItems() {
    setSelectedIds(visibleItems.map((item) => item.id));
  }

  function toggleSelectAllItems() {
    setSelectedIds(allSelected ? [] : visibleItems.map((item) => item.id));
  }

  function toggleSort(field: SortField) {
    if (sortField === field) {
      setSortDirection((current) => (current === "asc" ? "desc" : "asc"));
      return;
    }
    setSortField(field);
    setSortDirection("asc");
  }

  useKeyboardShortcuts({
    enabled: true,
    onSelectAll: selectAllItems,
    onDelete: () => {
      if (selectedItems.length > 0) setDialog({ type: "delete", items: selectedItems });
    },
    onRename: () => {
      const item = selectedItems[0];
      if (!item) return;
      if (selectedItems.length === 1 && features.rename) {
        setInlineRenameId(item.id);
        return;
      }
      setDialog({ type: "rename", item });
    },
    onRefresh: refreshListing,
    onOpen: () => {
      if (selectedItems[0]) onOpen(selectedItems[0]);
    },
  });

  function onOpen(item: SharePointItem) {
    if (item.type === "folder") {
      setCrumbs((current) => [...current, { id: item.id, name: item.name }]);
      setSelectedIds([]);
      setQuery("");
      return;
    }
    if (item.webUrl) {
      window.open(item.webUrl, "_blank", "noopener,noreferrer");
    }
  }

  function handleItemDragStart(item: SharePointItem, event: React.DragEvent) {
    if (!features.dragDropMove || !features.move) return;
    const dragIds = selectedIds.includes(item.id)
      ? sortedItems.filter((entry) => selectedIds.includes(entry.id)).map((entry) => entry.id)
      : [item.id];
    event.dataTransfer.setData(DRAG_MIME, JSON.stringify(dragIds));
    event.dataTransfer.effectAllowed = "move";
  }

  async function moveItemsToFolder(itemIds: string[], destinationParentId: string) {
    const movable = itemIds.filter((id) => id !== destinationParentId);
    if (movable.length === 0) return;
    setActionError(undefined);
    try {
      for (const itemId of movable) {
        await moveItem.mutateAsync({ itemId, destinationParentId });
      }
      setSelectedIds([]);
      reportSuccess(messages.operationSuccess);
    } catch (caught) {
      reportError(caught);
    }
  }

  function handleBreadcrumbDrop(destinationParentId: string, event: React.DragEvent) {
    event.preventDefault();
    setBreadcrumbDropTargetId(null);
    if (!features.dragDropMove || !features.move) return;
    const raw = event.dataTransfer.getData(DRAG_MIME);
    if (!raw) return;
    try {
      const itemIds = JSON.parse(raw) as string[];
      if (!Array.isArray(itemIds) || itemIds.length === 0) return;
      void moveItemsToFolder(itemIds, destinationParentId);
    } catch {
      // Ignore malformed drag payload.
    }
  }

  function openInSharePoint(item: SharePointItem) {
    if (item.webUrl) {
      window.open(item.webUrl, "_blank", "noopener,noreferrer");
    }
  }

  async function bulkDownload(targetItems: SharePointItem[]) {
    setActionError(undefined);
    for (const item of targetItems.filter((entry) => entry.type === "file")) {
      await downloadFile.mutateAsync(item);
    }
  }

  function onSelect(item: SharePointItem, additive: boolean, range: boolean) {
    const index = visibleItems.findIndex((entry) => entry.id === item.id);
    setSelectedIds((current) => {
      if (range && lastSelectedIndexRef.current >= 0 && index >= 0) {
        const start = Math.min(lastSelectedIndexRef.current, index);
        const end = Math.max(lastSelectedIndexRef.current, index);
        const rangeIds = visibleItems.slice(start, end + 1).map((entry) => entry.id);
        return Array.from(new Set([...current, ...rangeIds]));
      }
      if (additive) {
        return current.includes(item.id) ? current.filter((id) => id !== item.id) : [...current, item.id];
      }
      return [item.id];
    });
    if (index >= 0) lastSelectedIndexRef.current = index;
  }

  async function runAction(action: FileAction, item: SharePointItem) {
    setMenu(null);
    setActionError(undefined);
    setRestoreSuccess(false);
    try {
      if (action === "open") onOpen(item);
      if (action === "openInSharePoint") openInSharePoint(item);
      if (action === "download") await downloadFile.mutateAsync(item);
      if (action === "preview") {
        setDialog({ type: "preview", item });
        setPreview(await client.files.preview(item.id));
      }
      if (action === "rename") {
        if (features.rename) setInlineRenameId(item.id);
        else setDialog({ type: "rename", item });
      }
      if (action === "copy") setDialog({ type: "copy", items: [item] });
      if (action === "move") setDialog({ type: "move", items: [item] });
      if (action === "share") setDialog({ type: "share", item });
      if (action === "manageAccess") {
        setGrantSuccess(false);
        setDialog({ type: "access", item });
      }
      if (action === "properties") setDialog({ type: "properties", item });
      if (action === "activity") setDialog({ type: "activity", item });
      if (action === "delete") setDialog({ type: "delete", items: [item] });
      if (action === "versionHistory") {
        setRestoreSuccess(false);
        setDialog({ type: "versions", item });
      }
      if (action === "checkout") {
        await checkoutMutation.mutateAsync({ itemId: item.id, action: "checkout" });
      }
      if (action === "checkin") {
        setDialog({ type: "checkin", item });
      }
      if (action === "discardCheckout") {
        await checkoutMutation.mutateAsync({ itemId: item.id, action: "discardCheckout" });
      }
    } catch (caught) {
      reportError(caught);
    }
  }

  async function handleUpload(fileList: FileList, conflictBehavior: ConflictBehavior = "rename") {
    setActionError(undefined);
    uploadAbortRef.current = new AbortController();
    try {
      for (const file of Array.from(fileList)) {
        await uploadFile.mutateAsync({
          file,
          conflictBehavior,
          signal: uploadAbortRef.current.signal,
          onProgress: (progress) => setUploadProgress(progress.percent),
        });
      }
      setDialog(null);
      setUploadProgress(0);
      reportSuccess(messages.operationSuccess);
    } catch (caught) {
      if (uploadAbortRef.current.signal.aborted) {
        setDialog(null);
        setUploadProgress(0);
        return;
      }
      reportError(caught);
    } finally {
      uploadAbortRef.current = null;
    }
  }

  function cancelUpload() {
    uploadAbortRef.current?.abort();
  }

  async function runOnItems(
    targetItems: SharePointItem[],
    task: (item: SharePointItem) => Promise<unknown>,
  ) {
    setActionError(undefined);
    for (const item of targetItems) {
      await task(item);
    }
    setDialog(null);
    setSelectedIds([]);
  }

  async function handleFolderUpload(fileList: FileList, conflictBehavior: ConflictBehavior = "rename") {
    setActionError(undefined);
    uploadAbortRef.current = new AbortController();
    try {
      async function ensureFolderPath(baseParentId: string, folderNames: string[]): Promise<string> {
        let parentId = baseParentId;
        for (const folderName of folderNames) {
          const children = await client.folders.listChildren(parentId);
          const existing = children.items.find(
            (entry) => entry.type === "folder" && entry.name === folderName,
          );
          if (existing) {
            parentId = existing.id;
          } else {
            const created = await client.folders.create(parentId, folderName);
            parentId = created.id;
          }
        }
        return parentId;
      }

      const files = Array.from(fileList);
      for (const file of files) {
        const relativePath = file.webkitRelativePath || file.name;
        const parts = relativePath.split("/");
        const fileName = parts.pop() ?? file.name;
        const parentId = parts.length > 0 ? await ensureFolderPath(currentFolderId, parts) : currentFolderId;
        await client.upload.upload({
          parentId,
          fileName,
          content: file,
          conflictBehavior,
          signal: uploadAbortRef.current.signal,
        });
      }
      void childrenQuery.refetch();
      void infiniteChildrenQuery.refetch();
      setDialog(null);
    } catch (caught) {
      if (uploadAbortRef.current?.signal.aborted) {
        setDialog(null);
        return;
      }
      setActionError(getErrorMessage(caught, messages));
    } finally {
      uploadAbortRef.current = null;
    }
  }

  async function createOffice(kind: OfficeFileKind) {
    setActionError(undefined);
    try {
      await createOfficeFile.mutateAsync(kind);
    } catch (caught) {
      setActionError(getErrorMessage(caught, messages));
    }
  }

  const copyMoveItems = dialog?.type === "copy" || dialog?.type === "move" ? dialog.items : [];
  const allSelected = visibleItems.length > 0 && visibleItems.every((item) => selectedIds.includes(item.id));

  function handleSelectionAction(action: SelectionAction) {
    if (action === "download") void bulkDownload(selectedItems).catch((caught) => reportError(caught));
    if (action === "copy") setDialog({ type: "copy", items: selectedItems });
    if (action === "move") setDialog({ type: "move", items: selectedItems });
    if (action === "share" && selectedItems[0]) setDialog({ type: "share", item: selectedItems[0] });
    if (action === "rename" && selectedItems[0]) {
      if (features.rename) setInlineRenameId(selectedItems[0].id);
      else setDialog({ type: "rename", item: selectedItems[0] });
    }
    if (action === "preview" && selectedItems[0]) void runAction("preview", selectedItems[0]);
    if (action === "manageAccess" && selectedItems[0]) void runAction("manageAccess", selectedItems[0]);
    if (action === "bulkMetadata") setDialog({ type: "bulkMetadata", items: selectedItems });
    if (action === "delete") setDialog({ type: "delete", items: selectedItems });
  }

  const detailsItem = selectedItems.length === 1 ? selectedItems[0] : undefined;
  const detailsPermissionsQuery = usePermissions(
    detailsOpen && features.manageAccess && detailsItem ? detailsItem.id : undefined,
  );
  const detailsActivitiesQuery = useItemActivities(
    detailsOpen && features.activityLog && detailsItem ? detailsItem.id : undefined,
    detailsOpen && Boolean(detailsItem),
  );
  const chromeTitle = crumbs.length > 1
    ? crumbs[crumbs.length - 1]?.name ?? title ?? messages.files
    : title || messages.files;

  return (
    <div
      ref={dropRef}
      className={`spm-root spm-relative ${className ?? ""}`}
      style={{ borderColor: "var(--colorNeutralStroke2, #edebe9)" }}
      onDragEnter={(event) => {
        if (features.upload && event.dataTransfer.types.includes("Files")) setIsDragOver(true);
      }}
      onDragLeave={(event) => {
        if (event.currentTarget === event.target) setIsDragOver(false);
      }}
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => {
        event.preventDefault();
        setIsDragOver(false);
        if (features.upload && event.dataTransfer.files.length) {
          void handleUpload(event.dataTransfer.files).catch((caught) => reportError(caught));
        }
      }}
    >
      {isDragOver && features.upload ? (
        <div className="spm-drop-overlay">
          <span className="spm-text-sm spm-font-medium spm-text-sp-primary">{messages.dropFilesHere}</span>
        </div>
      ) : null}
      <input
        ref={folderUploadRef}
        type="file"
        className="spm-hidden"
        {...({ webkitdirectory: "", directory: "" } as React.InputHTMLAttributes<HTMLInputElement>)}
        multiple
        onChange={(event) => {
          if (event.target.files?.length) {
            void handleFolderUpload(event.target.files).catch((caught) => reportError(caught));
          }
          event.target.value = "";
        }}
      />

      <div className="spm-relative">
        <CommandBar
          messages={messages}
          features={features}
          title={chromeTitle}
          view={view}
          selectedItems={selectedItems}
          query={query}
          searchScope={searchScope}
          showFilters={showFilters}
          showColumnChooser={showColumnChooser}
          detailsOpen={detailsOpen}
          typeFilter={typeFilter}
          createOfficePending={createOfficeFile.isPending}
          onViewChange={setView}
          onNewFolder={() => setDialog({ type: "createFolder" })}
          onUpload={() => setDialog({ type: "upload" })}
          onUploadFolder={() => folderUploadRef.current?.click()}
          onCreateOffice={(kind) => void createOffice(kind)}
          onRefresh={refreshListing}
          onClearSelection={() => setSelectedIds([])}
          onSelectionAction={handleSelectionAction}
          onSearchChange={setQuery}
          onSearchScopeChange={setSearchScope}
          onToggleFilters={() => setShowFilters((current) => !current)}
          onToggleColumnChooser={() => setShowColumnChooser((current) => !current)}
          onToggleDetails={() => setDetailsOpen((current) => !current)}
          onTypeFilterChange={setTypeFilter}
        />
        <ColumnChooser
          open={showColumnChooser && (view === "list" || view === "compact")}
          messages={messages}
          settings={columnSettings}
          metadataColumns={features.metadata ? allMetadataColumns : []}
          onChange={updateColumnSettings}
          onClose={() => setShowColumnChooser(false)}
        />
      </div>

      <FilterPanel
        open={showFilters && features.globalSearch && searchScope === "library"}
        messages={messages}
        filters={searchFilters}
        onChange={setSearchFilters}
        onApply={() => {
          setAppliedFilters(searchFilters);
          setSearchExtraItems([]);
          setSearchNextLink(undefined);
        }}
        onClear={() => {
          setSearchFilters({});
          setAppliedFilters({});
          setSearchExtraItems([]);
          setSearchNextLink(undefined);
        }}
      />

      <nav
        className="spm-breadcrumb-bar"
        style={{
          background: tokens.colorNeutralBackground1,
        }}
      >
        <Breadcrumb aria-label="Folder navigation">
            {visibleCrumbs.map((crumb, visibleIndex) => {
              if (hiddenCrumbs.length > 0 && visibleIndex === 1) {
                return (
                  <BreadcrumbItem key="breadcrumb-overflow">
                    <Menu>
                      <MenuTrigger disableButtonEnhancement>
                        <BreadcrumbButton aria-label="More folders" icon={<MoreHorizontalRegular />} />
                      </MenuTrigger>
                      <MenuPopover>
                        <MenuList>
                          {hiddenCrumbs.map((hidden) => (
                            <MenuItem key={hidden.id} onClick={() => setCrumbs((current) => current.slice(0, current.findIndex((entry) => entry.id === hidden.id) + 1))}>
                              {hidden.name}
                            </MenuItem>
                          ))}
                        </MenuList>
                      </MenuPopover>
                    </Menu>
                  </BreadcrumbItem>
                );
              }
              const index = crumbs.findIndex((entry) => entry.id === crumb.id);
              const isLast = index === crumbs.length - 1;
              return (
                <BreadcrumbItem key={crumb.id}>
                  <BreadcrumbButton
                    current={isLast}
                    icon={index === 0 ? <FolderRegular /> : undefined}
                    onClick={() => !isLast && setCrumbs((current) => current.slice(0, index + 1))}
                    onDragOver={
                      features.dragDropMove && features.move
                        ? (event) => {
                            event.preventDefault();
                            setBreadcrumbDropTargetId(crumb.id);
                          }
                        : undefined
                    }
                    onDragLeave={features.dragDropMove ? () => setBreadcrumbDropTargetId(null) : undefined}
                    onDrop={
                      features.dragDropMove && features.move
                        ? (event) => handleBreadcrumbDrop(crumb.id, event)
                        : undefined
                    }
                    style={
                      breadcrumbDropTargetId === crumb.id
                        ? { background: tokens.colorBrandBackground2 }
                        : undefined
                    }
                  >
                    {crumb.name}
                  </BreadcrumbButton>
                </BreadcrumbItem>
              );
            })}
          </Breadcrumb>
        {!loading && !query.trim() ? (
          <Text size={200} className="spm-item-count" style={{ color: tokens.colorNeutralForeground3 }}>
            {visibleItems.length} {messages.itemCount}
          </Text>
        ) : null}
        {isSearching && !loading ? (
          <Text size={200} className="spm-item-count" style={{ color: tokens.colorNeutralForeground3 }}>
            {visibleItems.length} {messages.resultCount}
          </Text>
        ) : null}
      </nav>

      {actionError ? <ErrorBanner message={actionError} /> : null}
      {error ? (
        <ErrorBanner message={getErrorMessage(error, messages)} onRetry={refreshListing} />
      ) : null}

      {isSearching && searchNextLink ? (
        <div className="spm-border-b spm-border-sp-border spm-px-3 spm-py-2">
          <Button disabled={loadingMoreSearch} onClick={() => void loadMoreSearch()}>
            {loadingMoreSearch ? messages.loading : messages.loadMore}
          </Button>
        </div>
      ) : null}
      {!isSearching && folderNextLink ? (
        <div className="spm-border-b spm-border-sp-border spm-px-3 spm-py-2">
          <Button
            disabled={infiniteChildrenQuery.isFetchingNextPage}
            onClick={() => void infiniteChildrenQuery.fetchNextPage()}
          >
            {infiniteChildrenQuery.isFetchingNextPage ? messages.loading : messages.loadMore}
          </Button>
        </div>
      ) : null}

      <div className="spm-flex spm-min-h-0 spm-flex-1">
        <div className="spm-library-canvas spm-min-w-0 spm-flex-1 spm-overflow-auto">
        {loading ? <LibrarySkeleton compact={view === "compact"} /> : null}
        {!loading && visibleItems.length === 0 ? (
          <EmptyState
            messages={messages}
            features={features}
            isSearch={Boolean(query.trim())}
            onUpload={() => setDialog({ type: "upload" })}
            onNewFolder={() => setDialog({ type: "createFolder" })}
          />
        ) : null}
        {(view === "list" || view === "compact") && visibleItems.length > 0 ? (
          <FileList
            items={visibleItems}
            selectedIds={selectedIds}
            locale={locale}
            messages={messages}
            sortField={sortField}
            sortDirection={sortDirection}
            allSelected={allSelected}
            listColumns={features.metadata ? visibleMetadataColumns : []}
            columnVisibility={columnSettings}
            density={density === "compact" || view === "compact" ? "compact" : "normal"}
            inlineRenameId={inlineRenameId}
            onSort={toggleSort}
            onSortDirection={(field, direction) => {
              setSortField(field);
              setSortDirection(direction);
            }}
            onHideColumn={(column) => updateColumnSettings({ ...columnSettings, [column]: false })}
            onOpen={onOpen}
            onSelect={onSelect}
            onSelectAll={toggleSelectAllItems}
            onContextMenu={(item, x, y) => setMenu({ item, x, y })}
            onRowShare={features.share ? (item) => setDialog({ type: "share", item }) : undefined}
            onInlineRename={(item, name) => void handleInlineRename(item, name)}
            onCancelInlineRename={() => setInlineRenameId(undefined)}
            draggable={features.dragDropMove && features.move}
            onItemDragStart={handleItemDragStart}
            breadcrumbDropTargetId={breadcrumbDropTargetId}
            onFolderDrop={
              features.dragDropMove && features.move
                ? (folderId, event) => {
                    setBreadcrumbDropTargetId(folderId);
                    handleBreadcrumbDrop(folderId, event);
                  }
                : undefined
            }
            onAddColumn={() => setShowColumnChooser(true)}
          />
        ) : view === "grid" && visibleItems.length > 0 ? (
          <FileGrid
            items={visibleItems}
            selectedIds={selectedIds}
            onOpen={onOpen}
            onSelect={onSelect}
            onContextMenu={(item, x, y) => setMenu({ item, x, y })}
            draggable={features.dragDropMove && features.move}
            onItemDragStart={handleItemDragStart}
          />
        ) : null}
        </div>
        {detailsOpen && features.properties ? (
          detailsItem ? (
            <DetailsPane
              item={detailsItem}
              locale={locale}
              messages={messages}
              onClose={() => setDetailsOpen(false)}
              onOpenProperties={() => setDialog({ type: "properties", item: detailsItem })}
              onPreview={
                detailsItem.type === "file" && features.preview
                  ? () => void runAction("preview", detailsItem)
                  : undefined
              }
              onShare={features.share ? () => setDialog({ type: "share", item: detailsItem }) : undefined}
              permissions={detailsPermissionsQuery.data}
              permissionsLoading={detailsPermissionsQuery.isLoading}
              activities={detailsActivitiesQuery.data}
              activitiesLoading={detailsActivitiesQuery.isLoading}
              onOpenManageAccess={
                features.manageAccess ? () => setDialog({ type: "access", item: detailsItem }) : undefined
              }
              onOpenActivity={
                features.activityLog ? () => setDialog({ type: "activity", item: detailsItem }) : undefined
              }
            />
          ) : (
            <aside className="spm-details-pane">
              <div className="spm-details-header">
                <Text weight="semibold">{messages.details}</Text>
                <Button onClick={() => setDetailsOpen(false)}>{messages.cancel}</Button>
              </div>
              <p className="spm-p-4 spm-text-sm spm-text-sp-muted">{messages.selectToViewDetails}</p>
            </aside>
          )
        ) : null}
      </div>

      {menu ? (
        <ContextMenu
          item={menu.item}
          x={menu.x}
          y={menu.y}
          messages={messages}
          features={features}
          onClose={() => setMenu(null)}
          onAction={(action) => void runAction(action, menu.item)}
        />
      ) : null}

      <CreateFolderDialog
        open={dialog?.type === "createFolder"}
        messages={messages}
        pending={createFolder.isPending}
        onClose={() => setDialog(null)}
        onSubmit={(name) =>
          createFolder.mutate(name, {
            onSuccess: () => setDialog(null),
            onError: (caught) => setActionError(getErrorMessage(caught, messages)),
          })
        }
      />
      <RenameDialog
        key={dialog?.type === "rename" ? dialog.item.id : "rename"}
        item={dialog?.type === "rename" ? dialog.item : undefined}
        open={dialog?.type === "rename"}
        messages={messages}
        pending={renameItem.isPending}
        onClose={() => setDialog(null)}
        onSubmit={(name) => {
          if (dialog?.type !== "rename") return;
          renameItem.mutate(
            { itemId: dialog.item.id, name },
            {
              onSuccess: () => setDialog(null),
              onError: (caught) => setActionError(getErrorMessage(caught, messages)),
            },
          );
        }}
      />
      <DeleteDialog
        items={dialog?.type === "delete" ? dialog.items : []}
        open={dialog?.type === "delete"}
        messages={messages}
        pending={deleteItem.isPending}
        onClose={() => setDialog(null)}
        onConfirm={() => {
          if (dialog?.type !== "delete") return;
          void runOnItems(dialog.items, (item) => deleteItem.mutateAsync(item.id)).catch((caught) =>
            setActionError(getErrorMessage(caught, messages)),
          );
        }}
      />
      <CopyMoveDialog
        mode={dialog?.type === "move" ? "move" : "copy"}
        open={dialog?.type === "copy" || dialog?.type === "move"}
        messages={messages}
        rootId={rootId}
        rootName={messages.files}
        initialCrumbs={crumbs}
        excludeIds={copyMoveItems.map((item) => item.id)}
        singleItem={copyMoveItems.length === 1 ? copyMoveItems[0] : undefined}
        pending={copyItem.isPending || moveItem.isPending}
        copyProgress={copyProgress}
        showCopyProgress={features.copyProgress}
        onClose={() => {
          setCopyProgress(undefined);
          setDialog(null);
        }}
        onSubmit={(destinationParentId, newName) => {
          if (dialog?.type !== "copy" && dialog?.type !== "move") return;
          if (dialog.type === "move") {
            void runOnItems(dialog.items, (item) =>
              moveItem.mutateAsync({ itemId: item.id, destinationParentId, newName: newName || undefined }),
            )
              .then(() => reportSuccess(messages.operationSuccess))
              .catch((caught) => reportError(caught));
            return;
          }
          setCopyProgress(0);
          void (async () => {
            setActionError(undefined);
            try {
              for (const item of dialog.items) {
                await copyItem.mutateAsync({
                  itemId: item.id,
                  destinationParentId,
                  newName: newName || undefined,
                  onCopyProgress: features.copyProgress
                    ? (progress: CopyOperationProgress) => {
                        if (progress.percent !== undefined) setCopyProgress(progress.percent);
                        if (progress.phase === "failed") {
                          emitNotify({ type: "error", message: messages.copyFailed });
                        }
                      }
                    : undefined,
                });
              }
              setCopyProgress(undefined);
              setDialog(null);
              setSelectedIds([]);
              reportSuccess(messages.copyComplete);
            } catch (caught) {
              setCopyProgress(undefined);
              reportError(caught);
            }
          })();
        }}
      />
      <BulkMetadataDialog
        items={dialog?.type === "bulkMetadata" ? dialog.items : []}
        columns={listColumnsQuery.data ?? []}
        open={dialog?.type === "bulkMetadata"}
        messages={messages}
        pending={bulkUpdateListItemFields.isPending}
        onClose={() => setDialog(null)}
        onSubmit={(fields) => {
          if (dialog?.type !== "bulkMetadata") return;
          bulkUpdateListItemFields.mutate(
            { itemIds: dialog.items.map((item) => item.id), fields },
            {
              onSuccess: () => {
                setDialog(null);
                setSelectedIds([]);
                refreshListing();
                reportSuccess(messages.operationSuccess);
              },
              onError: (caught) => reportError(caught),
            },
          );
        }}
      />
      <CheckinDialog
        item={dialog?.type === "checkin" ? dialog.item : undefined}
        open={dialog?.type === "checkin"}
        messages={messages}
        pending={checkoutMutation.isPending}
        onClose={() => setDialog(null)}
        onSubmit={(comment) => {
          if (dialog?.type !== "checkin") return;
          checkoutMutation.mutate(
            { itemId: dialog.item.id, action: "checkin", comment },
            {
              onSuccess: () => setDialog(null),
              onError: (caught) => setActionError(getErrorMessage(caught, messages)),
            },
          );
        }}
      />
      <UploadDialog
        open={dialog?.type === "upload"}
        messages={messages}
        progress={uploadProgress}
        pending={uploadFile.isPending}
        onClose={() => setDialog(null)}
        onCancel={cancelUpload}
        onUpload={(files, conflictBehavior) =>
          void handleUpload(files, conflictBehavior).catch((caught) => setActionError(getErrorMessage(caught, messages)))
        }
      />
      <ShareDialog
        item={dialog?.type === "share" ? dialog.item : undefined}
        open={dialog?.type === "share"}
        messages={messages}
        pending={invite.isPending || createLink.isPending}
        error={actionError}
        onClose={() => {
          createLink.reset();
          setDialog(null);
        }}
        onInvite={(recipients, role, message, notify) => {
          invite.mutate(
            { recipients, role, message, sendInvitation: notify },
            {
              onSuccess: () => setDialog(null),
              onError: (caught) => setActionError(getErrorMessage(caught, messages)),
            },
          );
        }}
        createdLinkUrl={createLink.data?.link?.webUrl}
        onCreateLink={(scope, type, expirationDateTime) => {
          createLink.mutate(
            { scope, type, expirationDateTime },
            {
              onSuccess: async (permission) => {
                const url = permission.link?.webUrl;
                if (!url) return;
                try {
                  await navigator.clipboard.writeText(url);
                } catch {
                  // URL is still shown in the dialog.
                }
              },
              onError: (caught) => setActionError(getErrorMessage(caught, messages)),
            },
          );
        }}
      />
      <ManageAccessDialog
        item={dialog?.type === "access" ? dialog.item : undefined}
        open={dialog?.type === "access"}
        messages={messages}
        locale={locale}
        permissions={permissionsQuery.data ?? []}
        loading={permissionsQuery.isLoading}
        error={permissionsQuery.error ? getErrorMessage(permissionsQuery.error, messages) : undefined}
        grantPending={accessInvite.isPending}
        grantSuccess={grantSuccess}
        onClose={() => {
          setGrantSuccess(false);
          setDialog(null);
        }}
        onRemove={(permissionId) =>
          removePermission.mutate(permissionId, {
            onError: (caught) => setActionError(getErrorMessage(caught, messages)),
          })
        }
        onChangeRole={(permissionId, roles) =>
          updatePermission.mutate(
            { permissionId, roles },
            {
              onError: (caught) => setActionError(getErrorMessage(caught, messages)),
            },
          )
        }
        onCopyLink={(url) => void navigator.clipboard.writeText(url)}
        onGrant={(recipients, role) => {
          accessInvite.mutate(
            { recipients, role, sendInvitation: true },
            {
              onSuccess: () => setGrantSuccess(true),
              onError: (caught) => setActionError(getErrorMessage(caught, messages)),
            },
          );
        }}
      />
      <PropertiesDialog
        item={propertiesQuery.data ?? (dialog?.type === "properties" ? dialog.item : undefined)}
        open={dialog?.type === "properties"}
        messages={messages}
        locale={locale}
        loading={propertiesQuery.isLoading || listItemFieldsQuery.isLoading}
        listColumns={listColumnsQuery.data}
        listItemFields={listItemFieldsQuery.data}
        metadataPending={updateListItemFields.isPending}
        onClose={() => setDialog(null)}
        onSaveMetadata={
          features.metadata
            ? (fields) => {
                updateListItemFields.mutate(fields, {
                  onSuccess: () => {
                    void listItemFieldsQuery.refetch();
                    refreshListing();
                  },
                  onError: (caught) => setActionError(getErrorMessage(caught, messages)),
                });
              }
            : undefined
        }
      />
      <ActivityDialog
        open={dialog?.type === "activity"}
        title={
          dialog?.type === "activity"
            ? `${messages.activityLog}: ${dialog.item.name}`
            : messages.activityLog
        }
        activities={activitiesQuery.data ?? []}
        locale={locale}
        messages={messages}
        loading={activitiesQuery.isLoading}
        error={activitiesQuery.error ? getErrorMessage(activitiesQuery.error, messages) : undefined}
        onClose={() => setDialog(null)}
      />
      <PreviewDialog
        item={dialog?.type === "preview" ? dialog.item : undefined}
        open={dialog?.type === "preview"}
        preview={preview}
        messages={messages}
        onClose={() => setDialog(null)}
        onPrint={() => {
          if (preview?.getUrl) window.open(preview.getUrl, "_blank", "noopener,noreferrer");
        }}
      />
      <VersionHistoryDialog
        item={dialog?.type === "versions" ? dialog.item : undefined}
        open={dialog?.type === "versions"}
        versions={versionsQuery.data ?? []}
        locale={locale}
        messages={messages}
        loading={versionsQuery.isLoading}
        error={versionsQuery.error ? getErrorMessage(versionsQuery.error, messages) : undefined}
        restorePending={restoreVersion.isPending}
        restoreSuccess={restoreSuccess}
        downloadPending={downloadVersion.isPending}
        onClose={() => {
          setRestoreSuccess(false);
          setDialog(null);
        }}
        onDownload={(versionId) => {
          if (dialog?.type !== "versions") return;
          downloadVersion.mutate(
            { itemId: dialog.item.id, versionId },
            {
              onError: (caught) => setActionError(getErrorMessage(caught, messages)),
            },
          );
        }}
        onRestore={(versionId) => {
          if (dialog?.type !== "versions") return;
          restoreVersion.mutate(
            { itemId: dialog.item.id, versionId },
            {
              onSuccess: () => setRestoreSuccess(true),
              onError: (caught) => setActionError(getErrorMessage(caught, messages)),
            },
          );
        }}
      />
    </div>
  );
}
