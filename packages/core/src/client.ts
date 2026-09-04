import { resolveConfig } from "./config/resolve-config";
import { SharePointRestClient } from "./rest/client";
import { CreateDocumentService } from "./services/create-document";
import { CreateFolderService } from "./services/create-folder";
import { DeleteItemService } from "./services/delete";
import { FieldService } from "./services/fields";
import { FileService } from "./services/files";
import { FolderService } from "./services/folder";
import { resolveLibrary } from "./services/library";
import { PermissionService } from "./services/permissions";
import { SearchService } from "./services/search";
import { UploadFileService } from "./services/upload";
import type {
  LibraryContext,
  ResolvedSharePointConfig,
  SharePointConfig,
  TokenProvider,
} from "./types/models";

/**
 * Lớp ghép core: token + REST + cache thư viện + Folder / Field / Permission / Search / ghi.
 * UI không fetch trực tiếp — SharePoint đi qua this.rest / services.
 */
export class SharePointClient {
  readonly config: ResolvedSharePointConfig;
  readonly rest: SharePointRestClient;
  readonly folders: FolderService;
  readonly fields: FieldService;
  readonly files: FileService;
  readonly folderCreate: CreateFolderService;
  readonly documentCreate: CreateDocumentService;
  readonly fileUpload: UploadFileService;
  readonly itemDelete: DeleteItemService;
  readonly permissions: PermissionService;
  readonly search: SearchService;

  private libraryPromise?: Promise<LibraryContext>;

  constructor(config: SharePointConfig, fetchImpl?: typeof fetch) {
    this.config = resolveConfig(config);
    this.config.locale = this.config.locale || "vi-VN";
    this.rest = new SharePointRestClient({
      siteUrl: this.config.siteUrl,
      tokenProvider: this.config.tokenProvider,
      scopes: this.config.scopes,
      fetchImpl,
    });
    this.fields = new FieldService(this.rest, () => this.getLibrary(), this.config.locale);
    this.folders = new FolderService(this.rest, () => this.getLibrary(), this.fields);
    this.files = new FileService(this.rest, () => this.getLibrary(), this.folders);
    this.folderCreate = new CreateFolderService(this.rest, () => this.getLibrary());
    this.documentCreate = new CreateDocumentService(this.rest, () => this.getLibrary());
    this.fileUpload = new UploadFileService(this.rest, () => this.getLibrary());
    this.itemDelete = new DeleteItemService(this.rest, () => this.getLibrary());
    this.permissions = new PermissionService(this.rest, () => this.getLibrary());
    this.search = new SearchService(this.rest, () => this.getLibrary());
  }

  get tokenProvider(): TokenProvider {
    return this.config.tokenProvider;
  }

  get cacheScope(): string {
    return this.config.libraryName ?? "default";
  }

  /** Đổi locale runtime; fields sẽ được GET lại theo Accept-Language mới. */
  setLocale(locale?: string): void {
    const nextLocale = locale || "vi-VN";
    if (this.config.locale === nextLocale) return;
    this.config.locale = nextLocale;
    this.fields.setLocale(nextLocale);
  }

  /** Tìm document library một lần rồi dùng lại cho listChildren. */
  async getLibrary(): Promise<LibraryContext> {
    this.libraryPromise ??= resolveLibrary(
      this.rest,
      this.config.libraryName ?? "",
    );
    return this.libraryPromise;
  }
}
