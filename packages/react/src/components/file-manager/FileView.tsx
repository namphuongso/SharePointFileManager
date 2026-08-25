import type { ReactNode } from "react";
import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import {
  Table,
  TableBody,
  TableCell,
  TableHeader,
  TableHeaderCell,
  TableRow,
  mergeClasses,
} from "@fluentui/react-components";
import type { FileListProps } from "../../types";
import { formatBytes, formatItemCount, formatRelativeDate } from "../../utils/format";
import { FileTypeIcon } from "./FileTypeIcon";
import { useFileManagerStyles } from "./useFileManagerStyles";

/** Tổng số con trực tiếp — Folder.ItemCount, không $select computed ItemChildCount. */
const ITEM_CHILD_COUNT = "ItemChildCount";
const PERSON_FIELDS = new Set(["Author", "Editor"]);

/** Bảng một cấp kiểu document library SharePoint. */
export function FileList({
  items,
  locale,
  messages,
  onOpenFolder,
  extraColumns = [],
  fixedTitles,
}: FileListProps) {
  const styles = useFileManagerStyles();

  return (
    <Table size="small" aria-label={messages.files} className={styles.table} noNativeElements={false}>
      <TableHeader>
        <TableRow className={styles.headerRow}>
          <TableHeaderCell className={mergeClasses(styles.headerCell, styles.nameCell)}>
            <span className={styles.headerTitle}>{fixedTitles?.name ?? messages.name}</span>
          </TableHeaderCell>
          <TableHeaderCell className={mergeClasses(styles.headerCell, styles.modifiedCell)}>
            <span className={styles.headerTitle}>{fixedTitles?.modified ?? messages.modified}</span>
          </TableHeaderCell>
          {extraColumns.map((col) => (
            <TableHeaderCell
              key={col.internalName}
              className={mergeClasses(styles.headerCell, styles.extraCell)}
            >
              <span className={styles.headerTitle}>{col.title}</span>
            </TableHeaderCell>
          ))}
          <TableHeaderCell className={mergeClasses(styles.headerCell, styles.sizeCell)}>
            <span className={styles.headerTitle}>{fixedTitles?.size ?? messages.size}</span>
          </TableHeaderCell>
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
            extraColumns={extraColumns}
          />
        ))}
      </TableBody>
    </Table>
  );
}

type FileRowProps = Omit<FileListProps, "items" | "fixedTitles"> & {
  item: SharePointItem;
};

function FileRow({ item, locale, messages, onOpenFolder, extraColumns = [] }: FileRowProps) {
  const styles = useFileManagerStyles();
  const folder = item.type === "folder";

  return (
    <TableRow
      className={mergeClasses(styles.row, folder && styles.rowFolder)}
      onClick={() => folder && onOpenFolder(item)}
    >
      <TableCell className={mergeClasses(styles.cell, styles.nameCell)}>
        <div className={styles.nameInner}>
          <FileTypeIcon item={item} size="sm" />
          <span className={mergeClasses(styles.nameText, folder && styles.nameFolder)}>
            {item.name}
          </span>
        </div>
      </TableCell>
      <TableCell className={mergeClasses(styles.cell, styles.modifiedCell)}>
        {formatRelativeDate(item.lastModifiedDateTime, locale)}
      </TableCell>
      {extraColumns.map((col) => (
        <TableCell key={col.internalName} className={mergeClasses(styles.cell, styles.extraCell)}>
          {renderExtraCell(item, col.internalName, locale, styles.personPill)}
        </TableCell>
      ))}
      <TableCell className={mergeClasses(styles.cell, styles.sizeCell)}>
        {folder
          ? formatItemCount(item.childItemCount, messages.itemCount)
          : formatBytes(item.size)}
      </TableCell>
    </TableRow>
  );
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
