import type { ReactNode } from "react";
import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import {
  Table,
  TableBody,
  TableCell,
  TableHeader,
  TableRow,
  mergeClasses,
} from "@fluentui/react-components";
import type { FileListColumn, FileListProps } from "../../types";
import { defaultColumnWidth, minColumnWidth } from "../../utils/columnLayout";
import { formatBytes, formatItemCount, formatRelativeDate } from "../../utils/format";
import { ColumnHeaderMenu } from "./ColumnHeaderMenu";
import { FileTypeIcon } from "./FileTypeIcon";
import { useFileManagerStyles } from "./useFileManagerStyles";

/** Tổng số con trực tiếp — Folder.ItemCount, không $select computed ItemChildCount. */
const ITEM_CHILD_COUNT = "ItemChildCount";
const PERSON_FIELDS = new Set(["Author", "Editor"]);

function cellClass(styles: ReturnType<typeof useFileManagerStyles>, kind: FileListColumn["kind"]): string {
  if (kind === "name") return styles.nameCell;
  if (kind === "modified") return styles.modifiedCell;
  if (kind === "size") return styles.sizeCell;
  return styles.extraCell;
}

/** Bảng một cấp kiểu document library SharePoint. */
export function FileList({
  items,
  locale,
  messages,
  onOpenFolder,
  columns,
  columnWidths,
  onColumnResize,
  onColumnResizeEnd,
  onColumnReorder,
  sort,
  onSort,
  extraColumnMenuGroups,
}: FileListProps) {
  const styles = useFileManagerStyles();
  const tableWidth = columns.reduce(
    (sum, col) => sum + (columnWidths[col.internalName] ?? defaultColumnWidth(col.internalName)),
    0,
  );

  return (
    <Table
      size="small"
      aria-label={messages.files}
      className={styles.table}
      style={{ minWidth: tableWidth }}
      noNativeElements={false}
    >
      <colgroup>
        {columns.map((col) => {
          const width = columnWidths[col.internalName] ?? defaultColumnWidth(col.internalName);
          return (
            <col
              key={col.internalName}
              style={{ width, minWidth: minColumnWidth(col.internalName) }}
            />
          );
        })}
      </colgroup>
      <TableHeader>
        <TableRow className={styles.headerRow}>
          {columns.map((col) => {
            const width = columnWidths[col.internalName] ?? defaultColumnWidth(col.internalName);
            return (
              <ColumnHeaderMenu
                key={col.internalName}
                title={col.title}
                field={col.internalName}
                typeAsString={col.typeAsString}
                sort={sort}
                onSort={onSort}
                messages={messages}
                className={mergeClasses(styles.headerCell, cellClass(styles, col.kind))}
                extraGroups={extraColumnMenuGroups}
                width={width}
                minWidth={minColumnWidth(col.internalName)}
                onResize={(next) => onColumnResize(col.internalName, next)}
                onResizeEnd={(next) => onColumnResizeEnd(col.internalName, next)}
                onReorder={onColumnReorder}
              />
            );
          })}
        </TableRow>
      </TableHeader>
      <TableBody>
        {items.map((item) => (
          <FileRow
            key={item.id}
            item={item}
            locale={locale}
            messages={messages}
            onOpenFolder={onOpenFolder}
            columns={columns}
          />
        ))}
      </TableBody>
    </Table>
  );
}

type FileRowProps = {
  item: SharePointItem;
  locale: string;
  messages: FileListProps["messages"];
  onOpenFolder: FileListProps["onOpenFolder"];
  columns: FileListColumn[];
};

function FileRow({ item, locale, messages, onOpenFolder, columns }: FileRowProps) {
  const styles = useFileManagerStyles();
  const folder = item.type === "folder";

  return (
    <TableRow
      className={mergeClasses(styles.row, folder && styles.rowFolder)}
      onClick={() => folder && onOpenFolder(item)}
    >
      {columns.map((col) => (
        <TableCell
          key={col.internalName}
          className={mergeClasses(styles.cell, cellClass(styles, col.kind))}
          style={{ minWidth: minColumnWidth(col.internalName) }}
        >
          {renderColumnCell(item, col, locale, messages, styles)}
        </TableCell>
      ))}
    </TableRow>
  );
}

function renderColumnCell(
  item: SharePointItem,
  column: FileListColumn,
  locale: string,
  messages: FileListProps["messages"],
  styles: ReturnType<typeof useFileManagerStyles>,
): ReactNode {
  if (column.kind === "name") {
    const folder = item.type === "folder";
    return (
      <div className={styles.nameInner}>
        <FileTypeIcon item={item} size="sm" />
        <span className={mergeClasses(styles.nameText, folder && styles.nameFolder)}>{item.name}</span>
      </div>
    );
  }
  if (column.kind === "modified") {
    return formatRelativeDate(item.lastModifiedDateTime, locale);
  }
  if (column.kind === "size") {
    return item.type === "folder"
      ? formatItemCount(item.childItemCount, messages.itemCount)
      : formatBytes(item.size);
  }
  return renderExtraCell(item, column.internalName, locale, styles.personPill);
}

function renderExtraCell(
  item: SharePointItem,
  internalName: string,
  locale: string,
  personClassName: string,
): ReactNode {
  if (internalName === ITEM_CHILD_COUNT) {
    if (item.type !== "folder" || item.childItemCount === undefined) return "—";
    return String(item.childItemCount);
  }
  const value = item.fields?.[internalName];
  if (PERSON_FIELDS.has(internalName) || isPersonValue(value)) {
    const name = personTitle(value);
    if (!name) return "—";
    return <span className={personClassName}>{name}</span>;
  }
  return formatFieldValue(value, internalName, locale);
}

function isPersonValue(value: unknown): boolean {
  return typeof value === "object" && value !== null && "Title" in value && "Id" in value;
}

function personTitle(value: unknown): string | undefined {
  if (typeof value === "object" && value !== null && "Title" in value) {
    const title = (value as { Title?: unknown }).Title;
    return typeof title === "string" && title.trim() ? title : undefined;
  }
  return undefined;
}

function formatFieldValue(value: unknown, internalName: string, locale: string): string {
  if (value === null || value === undefined || value === "") return "—";
  if (internalName === "File_x0020_Size") {
    const size = typeof value === "number" ? value : Number(value);
    return Number.isFinite(size) ? formatBytes(size) : "—";
  }
  if (internalName === "Created" || internalName === "Modified") {
    return formatRelativeDate(String(value), locale);
  }
  if (typeof value === "object" && value !== null && "Title" in value) {
    return String((value as { Title?: unknown }).Title ?? "—");
  }
  if (typeof value === "boolean") return value ? "✓" : "—";
  return String(value);
}
