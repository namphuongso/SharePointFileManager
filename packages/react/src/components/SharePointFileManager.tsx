import { useMemo, useRef, useState } from "react";
import type { PreviewInfo, SharePointConfig, SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import { SharePointProvider } from "../provider/SharePointProvider";
import { useSharePoint } from "../provider/context";
import {
  getErrorMessage,
  useCopyItem,
  useCreateFolder,
  useCreateLink,
  useDeleteItem,
  useDownloadFile,
  useFolderChildren,
  useInvite,
  useMoveItem,
  usePermissions,
  useRemovePermission,
  useRenameItem,
  useSearchItems,
  useUploadFile,
  useVersions,
} from "../hooks/hooks";
import type { FileAction } from "./ContextMenu";
import { ContextMenu } from "./ContextMenu";
import { FileGrid, FileList } from "./FileView";
import { CopyMoveDialog, CreateFolderDialog, DeleteDialog, RenameDialog } from "./dialogs/BasicDialogs";
import { ShareDialog, UploadDialog } from "./dialogs/ShareUploadDialogs";
import { ManageAccessDialog, PreviewDialog, VersionHistoryDialog } from "./dialogs/AccessDialogs";
import { Button, ErrorBanner } from "./ui";
import type { Messages } from "../i18n/messages";

export interface SharePointFileManagerProps {
  config: SharePointConfig;
  locale?: string;
  view?: "list" | "grid";
  className?: string;
  messages?: Partial<Messages>;
}

export function SharePointFileManager(props: SharePointFileManagerProps) {
  return (
    <SharePointProvider config={props.config} locale={props.locale} messages={props.messages}>
      <FileManagerShell initialView={props.view ?? "list"} className={props.className} />
    </SharePointProvider>
  );
}

interface Crumb {
  id: string;
  name: string;
}

function FileManagerShell({ initialView, className }: { initialView: "list" | "grid"; className?: string }) {
  const { client, locale, messages } = useSharePoint();
  const rootId = client.config.rootItemId;
  const [crumbs, setCrumbs] = useState<Crumb[]>([{ id: rootId, name: messages.files }]);
  const currentFolderId = crumbs[crumbs.length - 1]?.id ?? rootId;
  const [view, setView] = useState<"list" | "grid">(initialView);
  const [query, setQuery] = useState("");
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [menu, setMenu] = useState<{ item: SharePointItem; x: number; y: number } | null>(null);
  const [dialog, setDialog] = useState<
    | { type: "createFolder" }
    | { type: "upload" }
    | { type: "rename"; item: SharePointItem }
    | { type: "delete"; item: SharePointItem }
    | { type: "copy"; item: SharePointItem }
    | { type: "move"; item: SharePointItem }
    | { type: "share"; item: SharePointItem }
    | { type: "access"; item: SharePointItem }
    | { type: "preview"; item: SharePointItem }
    | { type: "versions"; item: SharePointItem }
    | null
  >(null);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [preview, setPreview] = useState<PreviewInfo | undefined>();
  const [actionError, setActionError] = useState<string | undefined>();
  const dropRef = useRef<HTMLDivElement>(null);

  const childrenQuery = useFolderChildren(currentFolderId);
  const searchQuery = useSearchItems(query, currentFolderId);
  const items = useMemo(() => {
    if (query.trim().length >= 2) return searchQuery.data?.items ?? [];
    return childrenQuery.data?.items ?? [];
  }, [childrenQuery.data, query, searchQuery.data]);

  const selectedItem = items.find((item) => item.id === selectedIds[0]) ?? menu?.item;
  const activeItem =
    dialog && "item" in dialog ? dialog.item : selectedItem;

  const createFolder = useCreateFolder(currentFolderId);
  const renameItem = useRenameItem(currentFolderId);
  const deleteItem = useDeleteItem(currentFolderId);
  const copyItem = useCopyItem(currentFolderId);
  const moveItem = useMoveItem(currentFolderId);
  const uploadFile = useUploadFile(currentFolderId);
  const downloadFile = useDownloadFile();
  const invite = useInvite(activeItem?.id ?? "");
  const createLink = useCreateLink(activeItem?.id ?? "");
  const removePermission = useRemovePermission(activeItem?.id ?? "");
  const permissionsQuery = usePermissions(dialog?.type === "access" ? dialog.item.id : undefined);
  const versionsQuery = useVersions(dialog?.type === "versions" ? dialog.item.id : undefined, dialog?.type === "versions");

  const features = client.config.features;
  const loading = query.trim().length >= 2 ? searchQuery.isLoading : childrenQuery.isLoading;
  const error = query.trim().length >= 2 ? searchQuery.error : childrenQuery.error;

  function onOpen(item: SharePointItem) {
    if (item.type === "folder") {
      setCrumbs((current) => [...current, { id: item.id, name: item.name }]);
      setSelectedIds([]);
      setQuery("");
      return;
    }
    if (features.openInSharePoint && item.webUrl) {
      window.open(item.webUrl, "_blank", "noopener,noreferrer");
    }
  }

  function onSelect(item: SharePointItem, additive: boolean) {
    setSelectedIds((current) => {
      if (additive) {
        return current.includes(item.id) ? current.filter((id) => id !== item.id) : [...current, item.id];
      }
      return [item.id];
    });
  }

  async function runAction(action: FileAction, item: SharePointItem) {
    setMenu(null);
    setActionError(undefined);
    try {
      if (action === "open") onOpen(item);
      if (action === "download") await downloadFile.mutateAsync(item);
      if (action === "preview") {
        setDialog({ type: "preview", item });
        setPreview(await client.files.preview(item.id));
      }
      if (action === "rename") setDialog({ type: "rename", item });
      if (action === "copy") setDialog({ type: "copy", item });
      if (action === "move") setDialog({ type: "move", item });
      if (action === "share") setDialog({ type: "share", item });
      if (action === "manageAccess") setDialog({ type: "access", item });
      if (action === "delete") setDialog({ type: "delete", item });
      if (action === "versionHistory") setDialog({ type: "versions", item });
    } catch (caught) {
      setActionError(getErrorMessage(caught, messages));
    }
  }

  async function handleUpload(fileList: FileList) {
    setActionError(undefined);
    for (const file of Array.from(fileList)) {
      await uploadFile.mutateAsync({
        file,
        onProgress: (progress) => setUploadProgress(progress.percent),
      });
    }
    setDialog(null);
    setUploadProgress(0);
  }

  return (
    <div
      ref={dropRef}
      className={`spm-root ${className ?? ""}`}
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => {
        event.preventDefault();
        if (features.upload && event.dataTransfer.files.length) {
          void handleUpload(event.dataTransfer.files).catch((caught) =>
            setActionError(getErrorMessage(caught, messages)),
          );
        }
      }}
    >
      <header className="spm-flex spm-flex-wrap spm-items-center spm-gap-2 spm-border-b spm-border-sp-border spm-p-3">
        {features.createFolder ? (
          <Button onClick={() => setDialog({ type: "createFolder" })}>{messages.newFolder}</Button>
        ) : null}
        {features.upload ? <Button onClick={() => setDialog({ type: "upload" })}>{messages.upload}</Button> : null}
        {features.search ? (
          <input
            value={query}
            placeholder={messages.searchPlaceholder}
            onChange={(event) => setQuery(event.target.value)}
            className="spm-min-w-[180px] spm-flex-1 spm-rounded-md spm-border spm-border-sp-border spm-px-3 spm-py-1.5 spm-text-sm"
          />
        ) : null}
        <Button onClick={() => setView("list")}>{messages.list}</Button>
        <Button onClick={() => setView("grid")}>{messages.grid}</Button>
      </header>

      <nav className="spm-flex spm-flex-wrap spm-items-center spm-gap-1 spm-border-b spm-border-sp-border spm-px-3 spm-py-2 spm-text-sm">
        {crumbs.map((crumb, index) => (
          <span key={crumb.id} className="spm-flex spm-items-center spm-gap-1">
            {index > 0 ? <span className="spm-text-sp-muted">/</span> : null}
            <button
              type="button"
              className="hover:spm-underline"
              onClick={() => setCrumbs((current) => current.slice(0, index + 1))}
            >
              {crumb.name}
            </button>
          </span>
        ))}
        {!loading && !query.trim() ? (
          <span className="spm-ml-auto spm-text-xs spm-text-sp-muted">{items.length} mục</span>
        ) : null}
      </nav>

      {actionError ? <ErrorBanner message={actionError} /> : null}
      {error ? (
        <ErrorBanner
          message={getErrorMessage(error, messages)}
          onRetry={() => {
            void childrenQuery.refetch();
          }}
        />
      ) : null}

      <div className="spm-flex-1 spm-overflow-auto">
        {loading ? <div className="spm-p-6 spm-text-sm spm-text-sp-muted">{messages.loading}</div> : null}
        {!loading && items.length === 0 ? (
          <div className="spm-p-6 spm-text-sm spm-text-sp-muted">
            {query.trim() ? messages.noResults : messages.empty}
          </div>
        ) : null}
        {view === "list" ? (
          <FileList
            items={items}
            selectedIds={selectedIds}
            locale={locale}
            messages={messages}
            onOpen={onOpen}
            onSelect={onSelect}
            onContextMenu={(item, x, y) => setMenu({ item, x, y })}
          />
        ) : (
          <FileGrid
            items={items}
            selectedIds={selectedIds}
            onOpen={onOpen}
            onSelect={onSelect}
            onContextMenu={(item, x, y) => setMenu({ item, x, y })}
          />
        )}
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
        item={dialog?.type === "delete" ? dialog.item : undefined}
        open={dialog?.type === "delete"}
        messages={messages}
        pending={deleteItem.isPending}
        onClose={() => setDialog(null)}
        onConfirm={() => {
          if (dialog?.type !== "delete") return;
          deleteItem.mutate(dialog.item.id, {
            onSuccess: () => setDialog(null),
            onError: (caught) => setActionError(getErrorMessage(caught, messages)),
          });
        }}
      />
      <CopyMoveDialog
        mode={dialog?.type === "move" ? "move" : "copy"}
        open={dialog?.type === "copy" || dialog?.type === "move"}
        messages={messages}
        currentFolderId={currentFolderId}
        pending={copyItem.isPending || moveItem.isPending}
        onClose={() => setDialog(null)}
        onSubmit={(destinationParentId) => {
          if (dialog?.type !== "copy" && dialog?.type !== "move") return;
          const mutate = dialog.type === "copy" ? copyItem.mutate : moveItem.mutate;
          mutate(
            { itemId: dialog.item.id, destinationParentId },
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
        onUpload={(files) => void handleUpload(files).catch((caught) => setActionError(getErrorMessage(caught, messages)))}
      />
      <ShareDialog
        item={dialog?.type === "share" ? dialog.item : undefined}
        open={dialog?.type === "share"}
        messages={messages}
        pending={invite.isPending || createLink.isPending}
        error={actionError}
        onClose={() => setDialog(null)}
        onInvite={(email, role, message, notify) => {
          invite.mutate(
            { recipients: [{ email }], role, message, sendInvitation: notify },
            {
              onSuccess: () => setDialog(null),
              onError: (caught) => setActionError(getErrorMessage(caught, messages)),
            },
          );
        }}
        onCreateLink={(scope, type, expirationDateTime) => {
          createLink.mutate(
            { scope, type, expirationDateTime },
            {
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
        onClose={() => setDialog(null)}
        onRemove={(permissionId) =>
          removePermission.mutate(permissionId, {
            onError: (caught) => setActionError(getErrorMessage(caught, messages)),
          })
        }
        onCopyLink={(url) => void navigator.clipboard.writeText(url)}
      />
      <PreviewDialog
        item={dialog?.type === "preview" ? dialog.item : undefined}
        open={dialog?.type === "preview"}
        preview={preview}
        messages={messages}
        onClose={() => setDialog(null)}
      />
      <VersionHistoryDialog
        item={dialog?.type === "versions" ? dialog.item : undefined}
        open={dialog?.type === "versions"}
        versions={versionsQuery.data ?? []}
        locale={locale}
        messages={messages}
        onClose={() => setDialog(null)}
        onRestore={(versionId) => {
          if (dialog?.type !== "versions") return;
          void client.files.restoreVersion(dialog.item.id, versionId).catch((caught) =>
            setActionError(getErrorMessage(caught, messages)),
          );
        }}
      />
    </div>
  );
}
