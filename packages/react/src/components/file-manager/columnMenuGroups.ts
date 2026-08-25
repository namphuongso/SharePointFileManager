import { isSortableLibraryField } from "@namphuongso/sharepoint-file-manager-core";
import type { ListSortDirection } from "@namphuongso/sharepoint-file-manager-core";
import type { ColumnMenuGroup, Messages } from "../../types";

const NUMBER_TYPES = new Set(["number", "integer", "counter", "currency"]);

/** Menu header cột: hai hướng sort. extraGroups để host thêm mục sau. */
export function columnHeaderMenuGroups(options: {
  field: string;
  typeAsString?: string;
  sortDirection?: ListSortDirection;
  onSort?: (direction: ListSortDirection) => void;
  messages: Messages;
  extraGroups?: ColumnMenuGroup[];
}): ColumnMenuGroup[] {
  const { field, typeAsString, sortDirection, onSort, messages, extraGroups } = options;
  const groups: ColumnMenuGroup[] = [];

  if (onSort && isSortableLibraryField(field, typeAsString)) {
    const labels = sortLabels(field, typeAsString, messages);
    groups.push({
      id: "sort",
      items: [
        {
          id: "asc",
          label: labels.asc,
          checked: sortDirection === "asc",
          onClick: () => onSort("asc"),
        },
        {
          id: "desc",
          label: labels.desc,
          checked: sortDirection === "desc",
          onClick: () => onSort("desc"),
        },
      ],
    });
  }

  if (extraGroups?.length) groups.push(...extraGroups);
  return groups;
}

function sortLabels(
  field: string,
  typeAsString: string | undefined,
  messages: Messages,
): { asc: string; desc: string } {
  const type = typeAsString?.toLowerCase();
  if (field === "Modified" || field === "Created" || type === "datetime") {
    return { asc: messages.sortOldest, desc: messages.sortNewest };
  }
  if (field === "File_x0020_Size" || (type && NUMBER_TYPES.has(type))) {
    return { asc: messages.sortSmallest, desc: messages.sortLargest };
  }
  return { asc: messages.sortAZ, desc: messages.sortZA };
}
