import type {
  ListChildrenSort,
  ListSortDirection,
  SharePointConfig,
  SharePointItem,
  SharePointLibraryTarget,
} from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "./messages";

export interface SharePointFileManagerProps extends SharePointLibraryTarget {
  config?: SharePointConfig;
  locale?: string;
  /** Mặc định bật menu VI/EN trên toolbar; truyền false để ẩn. */
  showLanguageSwitcher?: boolean;
  className?: string;
  title?: string;
  messages?: Partial<Messages>;
  theme?: "light" | "dark" | "system";
}

export interface FileBrowserProps {
  className?: string;
  title?: string;
  showLanguageSwitcher?: boolean;
}

export interface BreadcrumbCrumb {
  id: string;
  name: string;
}

export type FileListColumnKind = "name" | "modified" | "size" | "extra";

export interface FileListColumn {
  internalName: string;
  title: string;
  typeAsString?: string;
  kind: FileListColumnKind;
}

export interface FileListProps {
  items: SharePointItem[];
  locale: string;
  messages: Messages;
  onOpenFolder: (item: SharePointItem) => void;
  /** Bấm dòng file — mở preview (?web=1) qua useOpenItem; bỏ trống để chỉ đọc. */
  onOpenFile?: (item: SharePointItem) => void;
  /** Tải xuống từ menu ⋯ / chuột phải. */
  onDownloadFile?: (item: SharePointItem) => void;
  /** Mở dialog xác nhận xóa từ menu ⋯ / chuột phải. */
  onDeleteFile?: (item: SharePointItem) => void;
  /** Chuột phải trên dòng — menu Mở / Tải xuống / Xóa; bỏ trống thì chỉ chặn menu mặc định. */
  onItemContextMenu?: (item: SharePointItem, position: { x: number; y: number }) => void;
  /** Đang mở / tải / xóa — khoá nút ⋯. */
  itemActionBusy?: boolean;
  columns: FileListColumn[];
  columnWidths: Record<string, number>;
  onColumnResize: (field: string, width: number) => void;
  onColumnResizeEnd: (widths: Record<string, number>) => void;
  onColumnReorder: (fromField: string, toField: string, place: "before" | "after") => void;
  sort?: ListChildrenSort;
  onSort?: (field: string, direction: ListSortDirection, typeAsString?: string) => void;
  /** Mặc định isSortableLibraryField; Search dùng isSortableSearchField (Size được). */
  isSortable?: (field: string, typeAsString?: string) => boolean;
  /** Nhóm menu thêm — nối sau sort. */
  extraColumnMenuGroups?: ColumnMenuGroup[];
}

/** Một dòng trong menu header cột; `submenu` để thêm cấp sau này. */
export interface ColumnMenuItem {
  id: string;
  label: string;
  disabled?: boolean;
  checked?: boolean;
  submenu?: ColumnMenuItem[];
  onClick?: () => void;
}

export interface ColumnMenuGroup {
  id: string;
  items: ColumnMenuItem[];
}

export interface ColumnHeaderMenuProps {
  title: string;
  field: string;
  typeAsString?: string;
  sort?: ListChildrenSort;
  onSort?: (field: string, direction: ListSortDirection, typeAsString?: string) => void;
  messages: Messages;
  className: string;
  isSortable?: (field: string, typeAsString?: string) => boolean;
  extraGroups?: ColumnMenuGroup[];
  width: number;
  minWidth: number;
  onResize: (width: number) => void;
  onResizeEnd: (width: number) => void;
  onReorder: (fromField: string, toField: string, place: "before" | "after") => void;
}

export interface EmptyStateProps {
  messages: Messages;
  /** Ghi đè tiêu đề / gợi ý (vd. tab Search). */
  title?: string;
  hint?: string;
}

export interface ForbiddenStateProps {
  messages: Messages;
}

export interface ErrorBannerProps {
  message: string;
  onRetry?: () => void;
  retryLabel?: string;
}

export type FileTypeIconSize = "sm" | "md" | "lg";

export interface FileTypeIconProps {
  item: SharePointItem;
  size?: FileTypeIconSize;
}
