import type { TokenProvider } from "./models";

export interface RestRequestOptions {
  path: string;
  /** URL đầy đủ (@odata.nextLink). Không ghép site/_api. */
  absoluteUrl?: string;
  method?: string;
  query?: Record<string, string | number | boolean | undefined>;
  headers?: Record<string, string>;
  signal?: AbortSignal;
}

export interface SharePointRestClientOptions {
  siteUrl: string;
  tokenProvider: TokenProvider;
  scopes: string[];
  fetchImpl?: typeof fetch;
}

export interface RestFile {
  UniqueId?: string;
  Name?: string;
  Length?: number | string;
  TimeLastModified?: string;
}

export interface RestFolder {
  UniqueId?: string;
  Name?: string;
  ServerRelativeUrl?: string;
  TimeLastModified?: string;
}

export interface RestList {
  Id?: string;
  Title?: string;
  EntityTypeName?: string;
  RootFolder?: { ServerRelativeUrl?: string; UniqueId?: string; Name?: string };
}

/** JSON list item (odata=nometadata). Không $select — đủ cột scalar + $expand. */
export interface RestListItem {
  Id?: number;
  GUID?: string;
  FSObjType?: number | string;
  FileSystemObjectType?: number;
  FileLeafRef?: string;
  FileDirRef?: string;
  Title?: string;
  File?: RestFile;
  Folder?: RestFolder;
  [key: string]: unknown;
}

export interface RestODataCollection<T> {
  value?: T[];
  "@odata.nextLink"?: string;
}

/** JSON SP.Field (odata=nometadata). Không $select — REST trả đủ property scalar. */
export interface RestField {
  Id?: string;
  Title?: string;
  InternalName?: string;
  EntityPropertyName?: string;
  TypeAsString?: string;
  FieldTypeKind?: number;
  Hidden?: boolean;
  ReadOnlyField?: boolean;
  Required?: boolean;
  Sortable?: boolean;
  Filterable?: boolean;
}

export interface RestErrorBody {
  error?: {
    code?: string;
    message?: { value?: string } | string;
  };
  "odata.error"?: {
    code?: string;
    message?: { value?: string };
  };
}

export interface MapRestErrorInput {
  status: number;
  body?: unknown;
  retryAfter?: string | null;
  fallbackMessage?: string;
}

/** Tùy chọn list một cấp folder (FolderService.listChildren). */
export interface ListChildrenOptions {
  /** Số dòng một trang — mặc định 30 như view SharePoint. */
  top?: number;
  /** Trang sau: GET nguyên @odata.nextLink. */
  nextLink?: string;
  signal?: AbortSignal;
}

/** Tùy chọn lấy schema cột library (FieldService.list). */
export interface ListFieldsOptions {
  signal?: AbortSignal;
}
