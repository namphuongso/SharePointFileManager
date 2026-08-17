import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../i18n/messages";
import { FileIcon, FolderIcon, formatBytes, formatDate } from "./ui";

export function FileList({
  items,
  selectedIds,
  locale,
  messages,
  onOpen,
  onSelect,
  onContextMenu,
}: {
  items: SharePointItem[];
  selectedIds: string[];
  locale: string;
  messages: Messages;
  onOpen: (item: SharePointItem) => void;
  onSelect: (item: SharePointItem, additive: boolean) => void;
  onContextMenu: (item: SharePointItem, x: number, y: number) => void;
}) {
  return (
    <table className="spm-w-full spm-text-left spm-text-sm">
      <thead className="spm-sticky spm-top-0 spm-bg-sp-surface spm-text-sp-muted">
        <tr>
          <th className="spm-px-4 spm-py-2 spm-font-medium">{messages.name}</th>
          <th className="spm-px-4 spm-py-2 spm-font-medium">{messages.modified}</th>
          <th className="spm-px-4 spm-py-2 spm-font-medium">{messages.size}</th>
        </tr>
      </thead>
      <tbody>
        {items.map((item) => {
          const selected = selectedIds.includes(item.id);
          return (
            <tr
              key={item.id}
              className={`spm-cursor-pointer hover:spm-bg-slate-50 ${selected ? "spm-bg-blue-50" : ""}`}
              onClick={(event) => onSelect(item, event.metaKey || event.ctrlKey)}
              onDoubleClick={() => onOpen(item)}
              onContextMenu={(event) => {
                event.preventDefault();
                onContextMenu(item, event.clientX, event.clientY);
              }}
            >
              <td className="spm-flex spm-items-center spm-gap-2 spm-px-4 spm-py-2">
                {item.type === "folder" ? <FolderIcon /> : <FileIcon />}
                <span>{item.name}</span>
              </td>
              <td className="spm-px-4 spm-py-2 spm-text-sp-muted">{formatDate(item.lastModifiedDateTime, locale)}</td>
              <td className="spm-px-4 spm-py-2 spm-text-sp-muted">
                {item.type === "folder" ? "—" : formatBytes(item.size)}
              </td>
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}

export function FileGrid({
  items,
  selectedIds,
  onOpen,
  onSelect,
  onContextMenu,
}: {
  items: SharePointItem[];
  selectedIds: string[];
  onOpen: (item: SharePointItem) => void;
  onSelect: (item: SharePointItem, additive: boolean) => void;
  onContextMenu: (item: SharePointItem, x: number, y: number) => void;
}) {
  return (
    <div className="spm-grid spm-grid-cols-2 spm-gap-3 spm-p-4 sm:spm-grid-cols-4 lg:spm-grid-cols-6">
      {items.map((item) => {
        const selected = selectedIds.includes(item.id);
        return (
          <button
            type="button"
            key={item.id}
            className={`spm-flex spm-flex-col spm-items-center spm-gap-2 spm-rounded-lg spm-border spm-p-3 spm-text-center ${
              selected ? "spm-border-sp-primary spm-bg-blue-50" : "spm-border-sp-border"
            }`}
            onClick={(event) => onSelect(item, event.metaKey || event.ctrlKey)}
            onDoubleClick={() => onOpen(item)}
            onContextMenu={(event) => {
              event.preventDefault();
              onContextMenu(item, event.clientX, event.clientY);
            }}
          >
            {item.thumbnailUrl ? (
              <img src={item.thumbnailUrl} alt="" className="spm-h-16 spm-w-16 spm-object-cover spm-rounded" />
            ) : item.type === "folder" ? (
              <FolderIcon />
            ) : (
              <FileIcon />
            )}
            <span className="spm-line-clamp-2 spm-w-full spm-text-xs">{item.name}</span>
          </button>
        );
      })}
    </div>
  );
}
