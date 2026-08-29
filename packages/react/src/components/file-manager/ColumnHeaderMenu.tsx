import { Fragment, useRef, useState, type DragEvent } from "react";
import {
  Menu,
  MenuDivider,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  TableHeaderCell,
  mergeClasses,
} from "@fluentui/react-components";
import {
  ArrowSortDownRegular,
  ArrowSortUpRegular,
  CheckmarkRegular,
  ChevronDownRegular,
} from "@fluentui/react-icons";
import type { ColumnHeaderMenuProps, ColumnMenuItem } from "../../types";
import { ColumnResizeHandle } from "./ColumnResizeHandle";
import { columnHeaderMenuGroups } from "./columnMenuGroups";
import { useFileManagerStyles } from "./useFileManagerStyles";

const DRAG_PREFIX = "sp-col:";

/** Menu option sát tên cột; header kéo thả đổi vị trí, mép phải kéo rộng. */
export function ColumnHeaderMenu({
  title,
  field,
  typeAsString,
  sort,
  onSort,
  messages,
  className,
  isSortable,
  extraGroups,
  width,
  minWidth,
  onResize,
  onResizeEnd,
  onReorder,
}: ColumnHeaderMenuProps) {
  const styles = useFileManagerStyles();
  const draggedRef = useRef(false);
  const dropSideRef = useRef<"before" | "after" | null>(null);
  const [dropSide, setDropSide] = useState<"before" | "after" | null>(null);
  const active = sort?.field === field;
  const direction = active ? sort.direction : undefined;
  const ariaSort = !direction ? "none" : direction === "asc" ? "ascending" : "descending";
  const groups = columnHeaderMenuGroups({
    field,
    typeAsString,
    sortDirection: direction,
    onSort: onSort ? (next) => onSort(field, next, typeAsString) : undefined,
    isSortable,
    messages,
    extraGroups,
  });

  function markDropSide(side: "before" | "after" | null) {
    dropSideRef.current = side;
    setDropSide(side);
  }

  function onDragStart(event: DragEvent<HTMLTableCellElement>) {
    draggedRef.current = true;
    event.dataTransfer.setData("text/plain", `${DRAG_PREFIX}${field}`);
    event.dataTransfer.effectAllowed = "move";
  }

  function onDragEnd() {
    markDropSide(null);
    window.setTimeout(() => {
      draggedRef.current = false;
    }, 0);
  }

  function onDragOver(event: DragEvent<HTMLTableCellElement>) {
    event.preventDefault();
    event.dataTransfer.dropEffect = "move";
    const rect = event.currentTarget.getBoundingClientRect();
    markDropSide(event.clientX < rect.left + rect.width / 2 ? "before" : "after");
  }

  function onDragLeave(event: DragEvent<HTMLTableCellElement>) {
    if (event.currentTarget.contains(event.relatedTarget as Node)) return;
    markDropSide(null);
  }

  function onDrop(event: DragEvent<HTMLTableCellElement>) {
    event.preventDefault();
    const raw = event.dataTransfer.getData("text/plain");
    const from = raw.startsWith(DRAG_PREFIX) ? raw.slice(DRAG_PREFIX.length) : "";
    const place = dropSideRef.current ?? "before";
    markDropSide(null);
    if (!from) return;
    onReorder(from, field, place);
  }

  return (
    <TableHeaderCell
      className={mergeClasses(
        className,
        dropSide === "before" && styles.headerCellDropBefore,
        dropSide === "after" && styles.headerCellDropAfter,
      )}
      style={{ minWidth }}
      aria-sort={groups.length > 0 ? ariaSort : undefined}
      draggable
      onDragStart={onDragStart}
      onDragEnd={onDragEnd}
      onDragOver={onDragOver}
      onDragLeave={onDragLeave}
      onDrop={onDrop}
    >
      {groups.length === 0 ? (
        <span className={styles.headerTitle}>{title}</span>
      ) : (
        <Menu positioning="below-start">
          <MenuTrigger disableButtonEnhancement>
            <div
              role="button"
              tabIndex={0}
              className={mergeClasses(
                styles.headerMenuButton,
                direction && styles.headerMenuButtonSorted,
              )}
              aria-label={
                !direction
                  ? title
                  : `${title}, ${direction === "asc" ? messages.sortAscending : messages.sortDescending}`
              }
              onClick={(event) => {
                if (draggedRef.current) event.preventDefault();
              }}
            >
              <span className={styles.headerTitle}>{title}</span>
              {direction === "asc" ? (
                <ArrowSortUpRegular fontSize={16} className={styles.headerSortGlyph} />
              ) : direction === "desc" ? (
                <ArrowSortDownRegular fontSize={16} className={styles.headerSortGlyph} />
              ) : (
                <ChevronDownRegular fontSize={12} className={styles.headerChevron} />
              )}
            </div>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              {groups.map((group, index) => (
                <Fragment key={group.id}>
                  {index > 0 ? <MenuDivider /> : null}
                  {group.items.map((item) => (
                    <ColumnMenuItemView key={item.id} item={item} />
                  ))}
                </Fragment>
              ))}
            </MenuList>
          </MenuPopover>
        </Menu>
      )}
      <ColumnResizeHandle
        width={width}
        minWidth={minWidth}
        label={messages.resizeColumn}
        onResize={onResize}
        onResizeEnd={onResizeEnd}
      />
    </TableHeaderCell>
  );
}

function ColumnMenuItemView({ item }: { item: ColumnMenuItem }) {
  if (item.submenu && item.submenu.length > 0) {
    return (
      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuItem disabled={item.disabled}>{item.label}</MenuItem>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            {item.submenu.map((child) => (
              <ColumnMenuItemView key={child.id} item={child} />
            ))}
          </MenuList>
        </MenuPopover>
      </Menu>
    );
  }

  return (
    <MenuItem
      disabled={item.disabled}
      icon={item.checked ? <CheckmarkRegular /> : undefined}
      onClick={item.onClick}
    >
      {item.label}
    </MenuItem>
  );
}
