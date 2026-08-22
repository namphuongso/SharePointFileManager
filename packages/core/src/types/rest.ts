import type { TokenProvider } from "./models";

export interface RestRequestOptions {
  path: string;
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
  Length?: number;
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
  top?: number;
  signal?: AbortSignal;
}
