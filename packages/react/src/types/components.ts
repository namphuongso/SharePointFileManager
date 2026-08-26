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
  columns: FileListColumn[];
  columnWidths: Record<string, number>;
  onColumnResize: (field: string, width: number) => void;
  onColumnResizeEnd: (field: string, width: number) => void;
  onColumnReorder: (fromField: string, toField: string, place: "before" | "after") => void;
  sort?: ListChildrenSort;
  onSort?: (field: string, direction: ListSortDirection, typeAsString?: string) => void;
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
  extraGroups?: ColumnMenuGroup[];
  width: number;
  minWidth: number;
  onResize: (width: number) => void;
  onResizeEnd: (width: number) => void;
  onReorder: (fromField: string, toField: string, place: "before" | "after") => void;
}

export interface EmptyStateProps {
  messages: Messages;
}

export interface ErrorBannerProps {
  message: string;
  onRetry?: () => void;
  retryLabel?: string;
}

export type FileKind =
  | "folder"
  | "word"
  | "excel"
  | "powerpoint"
  | "pdf"
  | "image"
  | "video"
  | "archive"
  | "generic";

export type FileTypeIconSize = "sm" | "md" | "lg";

export interface FileTypeIconProps {
  item: SharePointItem;
  size?: FileTypeIconSize;
}

export interface FolderIconProps {
  size: number;
}

export interface DocumentIconProps {
  color: string;
  accent: string;
  glyph: string;
  size: number;
}

export interface FileIconStyle {
  color: string;
  accent: string;
  glyph: string;
}
