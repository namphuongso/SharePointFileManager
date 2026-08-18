import { useEffect, useRef, useState } from "react";
import type { SharePointItem, SortDirection, SortField, ListColumn } from "@namphuongso/sharepoint-file-manager-core";
import {
  Button,
  Input,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Table,
  TableBody,
  TableCell,
  TableHeader,
  TableHeaderCell,
  TableRow,
  TableSelectionCell,
  tokens,
} from "@fluentui/react-components";
import {
  AddRegular,
  ArrowSortDownLinesRegular,
  ArrowSortUpLinesRegular,
  DismissRegular,
  LockClosedRegular,
  MoreHorizontalRegular,
  ShareRegular,
} from "@fluentui/react-icons";
import type { Messages } from "../i18n/messages";
import type { BuiltInColumnId, ColumnVisibilitySettings } from "../utils/column-settings";
import { FileTypeIcon } from "./FileTypeIcon";
import { formatBytes, formatRelativeDate } from "./ui";

export type ListDensity = "normal" | "compact";

function mapSortDirection(direction: SortDirection): "ascending" | "descending" {
  return direction === "asc" ? "ascending" : "descending";
}

export function FileList({
  items,
  selectedIds,
  locale,
  messages,
  sortField,
  sortDirection,
  allSelected,
  listColumns = [],
  columnVisibility,
  density = "normal",
  inlineRenameId,
  onSort,
  onSortDirection,
  onHideColumn,
  onOpen,
  onSelect,
  onSelectAll,
  onContextMenu,
  onRowShare,
  onInlineRename,
  onCancelInlineRename,
  draggable,
  onItemDragStart,
  onFolderDrop,
  breadcrumbDropTargetId,
  onAddColumn,
}: {
  items: SharePointItem[];
  selectedIds: string[];
  locale: string;
  messages: Messages;
  sortField: SortField;
  sortDirection: SortDirection;
  allSelected: boolean;
  listColumns?: ListColumn[];
  columnVisibility: ColumnVisibilitySettings;
  density?: ListDensity;
  inlineRenameId?: string;
  onSort: (field: SortField) => void;
  onSortDirection: (field: SortField, direction: SortDirection) => void;
  onHideColumn: (column: BuiltInColumnId) => void;
  onOpen: (item: SharePointItem) => void;
  onSelect: (item: SharePointItem, additive: boolean, range: boolean) => void;
  onSelectAll: () => void;
  onContextMenu: (item: SharePointItem, x: number, y: number) => void;
  onRowShare?: (item: SharePointItem) => void;
  onInlineRename?: (item: SharePointItem, name: string) => void;
  onCancelInlineRename?: () => void;
  draggable?: boolean;
  onItemDragStart?: (item: SharePointItem, event: React.DragEvent) => void;
  onFolderDrop?: (folderId: string, event: React.DragEvent) => void;
  breadcrumbDropTargetId?: string | null;
  onAddColumn?: () => void;
}) {
  const compact = density === "compact";

  return (
    <div className="spm-file-table-surface">
    <Table
      aria-label={messages.name}
      sortable
      className={`spm-file-table ${compact ? "spm-density-compact" : ""}`}
      style={{ background: tokens.colorNeutralBackground1 }}
    >
      <TableHeader>
        <TableRow>
          <TableSelectionCell
            checked={allSelected ? true : selectedIds.length > 0 ? "mixed" : false}
            onClick={onSelectAll}
            checkboxIndicator={{ "aria-label": messages.selectAll }}
          />
          <ColumnHeader
            label={messages.name}
            field="name"
            messages={messages}
            sortField={sortField}
            sortDirection={sortDirection}
            onSort={onSort}
            onSortDirection={onSortDirection}
            style={{ minWidth: 240 }}
          />
          {columnVisibility.modified ? (
            <ColumnHeader
              label={messages.modified}
              field="modified"
              columnId="modified"
              messages={messages}
              sortField={sortField}
              sortDirection={sortDirection}
              onSort={onSort}
              onSortDirection={onSortDirection}
              onHideColumn={onHideColumn}
            />
          ) : null}
          {columnVisibility.modifiedBy ? <TableHeaderCell>{messages.modifiedBy}</TableHeaderCell> : null}
          {columnVisibility.created ? (
            <ColumnHeader
              label={messages.created}
              field="created"
              columnId="created"
              messages={messages}
              sortField={sortField}
              sortDirection={sortDirection}
              onSort={onSort}
              onSortDirection={onSortDirection}
              onHideColumn={onHideColumn}
            />
          ) : null}
          {columnVisibility.createdBy ? <TableHeaderCell>{messages.createdBy}</TableHeaderCell> : null}
          {columnVisibility.size ? (
            <ColumnHeader
              label={messages.size}
              field="size"
              columnId="size"
              messages={messages}
              sortField={sortField}
              sortDirection={sortDirection}
              onSort={onSort}
              onSortDirection={onSortDirection}
              onHideColumn={onHideColumn}
            />
          ) : null}
          {listColumns.map((column) => (
            <TableHeaderCell key={column.id}>{column.displayName}</TableHeaderCell>
          ))}
          <TableHeaderCell style={{ width: 140 }}>
            <Button
              appearance="subtle"
              size="small"
              icon={<AddRegular />}
              aria-label={messages.addColumn}
              onClick={onAddColumn}
            >
              {messages.addColumn}
            </Button>
          </TableHeaderCell>
        </TableRow>
      </TableHeader>
      <TableBody>
        {items.map((item) => {
          const selected = selectedIds.includes(item.id);
          const isDropTarget = item.type === "folder" && breadcrumbDropTargetId === item.id;
          const isRenaming = inlineRenameId === item.id;
          return (
            <TableRow
              key={item.id}
              aria-selected={selected}
              className={`${selected ? "spm-row-selected" : "spm-row-hover"} ${isDropTarget ? "spm-row-drop-target" : ""}`}
              draggable={draggable}
              style={
                isDropTarget
                  ? { backgroundColor: tokens.colorBrandBackground2 }
                  : compact
                    ? { height: 32 }
                    : undefined
              }
              onClick={(event) => onSelect(item, event.metaKey || event.ctrlKey, event.shiftKey)}
              onDoubleClick={() => {
                if (!isRenaming) onOpen(item);
              }}
              onDragStart={(event) => onItemDragStart?.(item, event)}
              onDragOver={
                item.type === "folder" && onFolderDrop
                  ? (event) => {
                      event.preventDefault();
                    }
                  : undefined
              }
              onDrop={item.type === "folder" && onFolderDrop ? (event) => onFolderDrop(item.id, event) : undefined}
              onContextMenu={(event) => {
                event.preventDefault();
                onContextMenu(item, event.clientX, event.clientY);
              }}
            >
              <TableSelectionCell
                checked={selected}
                checkboxIndicator={{ "aria-label": item.name }}
                onClick={(event) => {
                  event.stopPropagation();
                  onSelect(item, true, event.shiftKey);
                }}
              />
              <TableCell className="spm-name-cell">
                <div className="spm-flex spm-min-w-0 spm-items-center spm-gap-2">
                  <FileTypeIcon item={item} size={compact ? "sm" : "md"} />
                  <InlineNameCell
                    item={item}
                    isRenaming={isRenaming}
                    onOpen={() => onOpen(item)}
                    onSubmit={(name) => onInlineRename?.(item, name)}
                    onCancel={() => onCancelInlineRename?.()}
                  />
                  {item.capabilities?.isCheckedOut ? (
                    <LockClosedRegular
                      style={{ color: tokens.colorPaletteDarkOrangeForeground1, flexShrink: 0 }}
                      title={messages.checkedOut}
                    />
                  ) : null}
                </div>
              </TableCell>
              {columnVisibility.modified ? (
                <TableCell>{formatRelativeDate(item.lastModifiedDateTime, locale)}</TableCell>
              ) : null}
              {columnVisibility.modifiedBy ? (
                <TableCell>{item.lastModifiedBy?.displayName ?? item.lastModifiedBy?.email ?? "—"}</TableCell>
              ) : null}
              {columnVisibility.created ? (
                <TableCell>{formatRelativeDate(item.createdDateTime, locale)}</TableCell>
              ) : null}
              {columnVisibility.createdBy ? (
                <TableCell>{item.createdBy?.displayName ?? item.createdBy?.email ?? "—"}</TableCell>
              ) : null}
              {columnVisibility.size ? (
                <TableCell>{item.type === "folder" ? "—" : formatBytes(item.size)}</TableCell>
              ) : null}
              {listColumns.map((column) => (
                <TableCell key={column.id}>
                  {formatMetadataValue(item.metadata?.[column.name], column, messages, locale)}
                </TableCell>
              ))}
              <TableCell>
                <div className="spm-row-actions">
                  {onRowShare && item.type === "file" ? (
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<ShareRegular />}
                      className="spm-row-action"
                      title={messages.share}
                      onClick={(event) => {
                        event.stopPropagation();
                        onRowShare(item);
                      }}
                    />
                  ) : null}
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<MoreHorizontalRegular />}
                    className="spm-row-action"
                    title={`${messages.moreActions}: ${item.name}`}
                    aria-label={`${messages.moreActions}: ${item.name}`}
                    onClick={(event) => {
                      event.stopPropagation();
                      onContextMenu(item, event.clientX, event.clientY);
                    }}
                  />
                </div>
              </TableCell>
            </TableRow>
          );
        })}
      </TableBody>
    </Table>
    </div>
  );
}

function ColumnHeader({
  label,
  field,
  columnId,
  messages,
  sortField,
  sortDirection,
  onSort,
  onSortDirection,
  onHideColumn,
  style,
}: {
  label: string;
  field: SortField;
  columnId?: BuiltInColumnId;
  messages: Messages;
  sortField: SortField;
  sortDirection: SortDirection;
  onSort: (field: SortField) => void;
  onSortDirection: (field: SortField, direction: SortDirection) => void;
  onHideColumn?: (column: BuiltInColumnId) => void;
  style?: React.CSSProperties;
}) {
  const headerRef = useRef<HTMLDivElement>(null);
  const [width, setWidth] = useState<number>();

  function startResize(event: React.PointerEvent<HTMLSpanElement>) {
    event.preventDefault();
    event.stopPropagation();
    const startX = event.clientX;
    const startWidth = headerRef.current?.closest("th")?.getBoundingClientRect().width ?? 160;
    const pointerId = event.pointerId;
    event.currentTarget.setPointerCapture(pointerId);

    function handleMove(moveEvent: PointerEvent) {
      setWidth(Math.max(96, startWidth + moveEvent.clientX - startX));
    }
    function handleUp() {
      window.removeEventListener("pointermove", handleMove);
      window.removeEventListener("pointerup", handleUp);
    }
    window.addEventListener("pointermove", handleMove);
    window.addEventListener("pointerup", handleUp, { once: true });
  }

  return (
    <TableHeaderCell
      sortDirection={sortField === field ? mapSortDirection(sortDirection) : undefined}
      sortable
      onClick={() => onSort(field)}
      className="spm-column-header"
      style={{ ...style, width, minWidth: width ?? style?.minWidth, maxWidth: width }}
    >
      <div ref={headerRef} className="spm-column-header-content">
        <span className="spm-column-header-label">{label}</span>
      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <Button
            appearance="subtle"
            size="small"
            icon={<MoreHorizontalRegular />}
            className="spm-column-menu-trigger"
            aria-label={`${messages.columns}: ${label}`}
            onClick={(event) => event.stopPropagation()}
          />
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuItem icon={<ArrowSortUpLinesRegular />} onClick={() => onSortDirection(field, "asc")}>
              {messages.sortAscending}
            </MenuItem>
            <MenuItem icon={<ArrowSortDownLinesRegular />} onClick={() => onSortDirection(field, "desc")}>
              {messages.sortDescending}
            </MenuItem>
            {columnId && onHideColumn ? (
              <MenuItem icon={<DismissRegular />} onClick={() => onHideColumn(columnId)}>
                {messages.hideColumn}
              </MenuItem>
            ) : null}
          </MenuList>
        </MenuPopover>
        </Menu>
      </div>
      <span className="spm-column-resize-handle" onPointerDown={startResize} aria-hidden="true" />
    </TableHeaderCell>
  );
}

function InlineNameCell({
  item,
  isRenaming,
  onOpen,
  onSubmit,
  onCancel,
}: {
  item: SharePointItem;
  isRenaming: boolean;
  onOpen: () => void;
  onSubmit: (name: string) => void;
  onCancel: () => void;
}) {
  const [draft, setDraft] = useState(item.name);

  useEffect(() => {
    if (isRenaming) setDraft(item.name);
  }, [isRenaming, item.name]);

  if (isRenaming) {
    return (
      <Input
        autoFocus
        value={draft}
        size="small"
        style={{ flex: 1, minWidth: 0 }}
        onClick={(event) => event.stopPropagation()}
        onChange={(_, data) => setDraft(data.value)}
        onKeyDown={(event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            const trimmed = draft.trim();
            if (trimmed && trimmed !== item.name) onSubmit(trimmed);
            else onCancel();
          }
          if (event.key === "Escape") {
            event.preventDefault();
            onCancel();
          }
        }}
        onBlur={() => {
          const trimmed = draft.trim();
          if (trimmed && trimmed !== item.name) onSubmit(trimmed);
          else onCancel();
        }}
      />
    );
  }

  return (
    <button
      type="button"
      className="spm-file-name-button spm-truncate"
      title={item.name}
      onClick={(event) => {
        event.stopPropagation();
        onOpen();
      }}
    >
      {item.name}
    </button>
  );
}

export function FileGrid({
  items,
  selectedIds,
  onOpen,
  onSelect,
  onContextMenu,
  draggable,
  onItemDragStart,
}: {
  items: SharePointItem[];
  selectedIds: string[];
  onOpen: (item: SharePointItem) => void;
  onSelect: (item: SharePointItem, additive: boolean, range: boolean) => void;
  onContextMenu: (item: SharePointItem, x: number, y: number) => void;
  draggable?: boolean;
  onItemDragStart?: (item: SharePointItem, event: React.DragEvent) => void;
}) {
  return (
    <div className="spm-grid-view">
      {items.map((item) => {
        const selected = selectedIds.includes(item.id);
        return (
          <button
            type="button"
            key={item.id}
            draggable={draggable}
            className={`spm-grid-item ${selected ? "selected" : ""}`}
            style={
              selected
                ? {
                    background: tokens.colorBrandBackground2,
                    borderColor: tokens.colorBrandStroke1,
                  }
                : undefined
            }
            onClick={(event) => onSelect(item, event.metaKey || event.ctrlKey, event.shiftKey)}
            onDoubleClick={() => onOpen(item)}
            onDragStart={(event) => onItemDragStart?.(item, event)}
            onContextMenu={(event) => {
              event.preventDefault();
              onContextMenu(item, event.clientX, event.clientY);
            }}
          >
            {item.thumbnailUrl ? (
              <img
                src={item.thumbnailUrl}
                alt=""
                className="spm-h-16 spm-w-16 spm-object-cover spm-rounded-sm"
              />
            ) : (
              <FileTypeIcon item={item} size="lg" />
            )}
            <span className="spm-grid-item-name">{item.name}</span>
          </button>
        );
      })}
    </div>
  );
}

function formatMetadataValue(
  value: string | number | boolean | null | undefined,
  column: ListColumn,
  messages: Messages,
  locale: string,
): string {
  if (value === null || value === undefined || value === "") return "—";
  if (typeof value === "boolean") return value ? messages.yes : messages.no;
  if (column.type === "dateTime" && typeof value === "string") {
    const date = new Date(value);
    if (!Number.isNaN(date.getTime())) {
      return new Intl.DateTimeFormat(locale, { dateStyle: "short" }).format(date);
    }
  }
  return String(value);
}
