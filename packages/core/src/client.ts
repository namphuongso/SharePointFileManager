import { resolveConfig } from "./config/resolve-config";
import { SharePointRestClient } from "./rest/client";
import { FieldService } from "./services/fields";
import { FolderService } from "./services/folder";
import { resolveLibrary } from "./services/library";
import type {
  LibraryContext,
  ResolvedSharePointConfig,
  SharePointConfig,
  TokenProvider,
} from "./types/models";

/**
 * Lớp ghép core: token + REST + cache thư viện + FolderService + FieldService.
 * UI không fetch trực tiếp — GET SharePoint đi qua this.rest / this.folders / this.fields.
 */
export class SharePointClient {
  readonly config: ResolvedSharePointConfig;
  readonly rest: SharePointRestClient;
  readonly folders: FolderService;
  readonly fields: FieldService;

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
