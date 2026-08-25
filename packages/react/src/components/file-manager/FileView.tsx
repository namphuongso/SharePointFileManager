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
import type { FileListProps } from "../../types";
import { formatBytes, formatItemCount, formatRelativeDate } from "../../utils/format";
import { ColumnHeaderMenu } from "./ColumnHeaderMenu";
import { FileTypeIcon } from "./FileTypeIcon";
import { useFileManagerStyles } from "./useFileManagerStyles";

/** Tổng số con trực tiếp — Folder.ItemCount, không $select computed ItemChildCount. */
const ITEM_CHILD_COUNT = "ItemChildCount";
const PERSON_FIELDS = new Set(["Author", "Editor"]);

const FIXED_SORT = {
  name: { field: "FileLeafRef", typeAsString: "Text" },
  modified: { field: "Modified", typeAsString: "DateTime" },
  size: { field: "File_x0020_Size", typeAsString: "Number" },
} as const;

/** Bảng một cấp kiểu document library SharePoint. */
export function FileList({
  items,
  locale,
  messages,
  onOpenFolder,
  extraColumns = [],
  fixedTitles,
  sort,
  onSort,
  extraColumnMenuGroups,
}: FileListProps) {
  const styles = useFileManagerStyles();

  return (
    <Table size="small" aria-label={messages.files} className={styles.table} noNativeElements={false}>
      <TableHeader>
        <TableRow className={styles.headerRow}>
          <ColumnHeaderMenu
            title={fixedTitles?.name ?? messages.name}
            field={FIXED_SORT.name.field}
            typeAsString={FIXED_SORT.name.typeAsString}
            sort={sort}
            onSort={onSort}
            messages={messages}
            className={mergeClasses(styles.headerCell, styles.nameCell)}
            extraGroups={extraColumnMenuGroups}
          />
          <ColumnHeaderMenu
            title={fixedTitles?.modified ?? messages.modified}
            field={FIXED_SORT.modified.field}
            typeAsString={FIXED_SORT.modified.typeAsString}
            sort={sort}
            onSort={onSort}
            messages={messages}
            className={mergeClasses(styles.headerCell, styles.modifiedCell)}
            extraGroups={extraColumnMenuGroups}
          />
          {extraColumns.map((col) => (
            <ColumnHeaderMenu
              key={col.internalName}
              title={col.title}
              field={col.internalName}
              typeAsString={col.typeAsString}
              sort={sort}
              onSort={onSort}
              messages={messages}
              className={mergeClasses(styles.headerCell, styles.extraCell)}
              extraGroups={extraColumnMenuGroups}
            />
          ))}
          <ColumnHeaderMenu
            title={fixedTitles?.size ?? messages.size}
            field={FIXED_SORT.size.field}
            typeAsString={FIXED_SORT.size.typeAsString}
            sort={sort}
            onSort={onSort}
            messages={messages}
            className={mergeClasses(styles.headerCell, styles.sizeCell)}
            extraGroups={extraColumnMenuGroups}
          />
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

type FileRowProps = Omit<
  FileListProps,
  "items" | "fixedTitles" | "sort" | "onSort" | "extraColumnMenuGroups"
> & {
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
