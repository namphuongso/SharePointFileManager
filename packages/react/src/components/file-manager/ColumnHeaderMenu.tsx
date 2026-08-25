import { Fragment } from "react";
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
import { columnHeaderMenuGroups } from "./columnMenuGroups";
import { useFileManagerStyles } from "./useFileManagerStyles";

/** Menu option sát tên cột; click lại hướng đang chọn thì bỏ sort. */
export function ColumnHeaderMenu({
  title,
  field,
  typeAsString,
  sort,
  onSort,
  messages,
  className,
  extraGroups,
}: ColumnHeaderMenuProps) {
  const styles = useFileManagerStyles();
  const active = sort?.field === field;
  const direction = active ? sort.direction : undefined;
  const ariaSort = !direction ? "none" : direction === "asc" ? "ascending" : "descending";
  const groups = columnHeaderMenuGroups({
    field,
    typeAsString,
    sortDirection: direction,
    onSort: onSort ? (next) => onSort(field, next, typeAsString) : undefined,
    messages,
    extraGroups,
  });

  if (groups.length === 0) {
    return (
      <TableHeaderCell className={className}>
        <span className={styles.headerTitle}>{title}</span>
      </TableHeaderCell>
    );
  }

  return (
    <TableHeaderCell className={className} aria-sort={ariaSort}>
      <Menu positioning="below-start">
        <MenuTrigger disableButtonEnhancement>
          <button
            type="button"
            className={mergeClasses(styles.headerMenuButton, direction && styles.headerMenuButtonSorted)}
            aria-label={
              !direction
                ? title
                : `${title}, ${direction === "asc" ? messages.sortAscending : messages.sortDescending}`
            }
          >
            <span className={styles.headerTitle}>{title}</span>
            {direction === "asc" ? (
              <ArrowSortUpRegular fontSize={16} className={styles.headerSortGlyph} />
            ) : direction === "desc" ? (
              <ArrowSortDownRegular fontSize={16} className={styles.headerSortGlyph} />
            ) : (
              <ChevronDownRegular fontSize={12} className={styles.headerChevron} />
            )}
          </button>
        </MenuTrigger>
        <MenuPopover className={styles.headerMenuPopover}>
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
