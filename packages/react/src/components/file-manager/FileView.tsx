import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import {
  Table,
  TableBody,
  TableCell,
  TableHeader,
  TableHeaderCell,
  TableRow,
  Text,
  tokens,
} from "@fluentui/react-components";
import type { FileListProps } from "../../types";
import { formatBytes, formatRelativeDate } from "../../utils/format";
import { FileTypeIcon } from "./FileTypeIcon";

/** Bảng một cấp: tên, sửa đổi, kích thước. Click folder thì gọi onOpenFolder. */

export function FileList({
  items,
  locale,
  messages,
  onOpenFolder,
}: FileListProps) {
  return (
    <Table size="small" aria-label={messages.files} style={{ minWidth: 640 }}>
      <TableHeader>
        <TableRow>
          <TableHeaderCell>{messages.name}</TableHeaderCell>
          <TableHeaderCell>{messages.modified}</TableHeaderCell>
          <TableHeaderCell>{messages.size}</TableHeaderCell>
        </TableRow>
      </TableHeader>
      <TableBody>
        {items.map((item) => (
          <TableRow
            key={item.id}
            onClick={() => item.type === "folder" && onOpenFolder(item)}
            style={{ cursor: item.type === "folder" ? "pointer" : "default" }}
          >
            <TableCell>
              <div style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0 }}>
                <FileTypeIcon item={item} size="sm" />
                <Text truncate wrap={false} weight={item.type === "folder" ? "semibold" : "regular"}>
                  {item.name}
                </Text>
              </div>
            </TableCell>
            <TableCell>
              <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                {formatRelativeDate(item.lastModifiedDateTime, locale)}
              </Text>
            </TableCell>
            <TableCell>
              <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                {item.type === "folder" ? "—" : formatBytes(item.size)}
              </Text>
            </TableCell>
          </TableRow>
        ))}
      </TableBody>
    </Table>
  );
}
