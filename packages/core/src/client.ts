import { resolveConfig } from "./config/resolve-config";
import { SharePointRestClient } from "./rest/client";
import { FolderService } from "./services/folder";
import { resolveLibrary } from "./services/library";
import type {
  LibraryContext,
  ResolvedSharePointConfig,
  SharePointConfig,
  TokenProvider,
} from "./types/models";

/**
 * Lớp ghép core: token + REST + cache thư viện + FolderService.
 * UI không fetch trực tiếp — mọi GET đi qua this.rest / this.folders.
 */
export class SharePointClient {
  readonly config: ResolvedSharePointConfig;
  readonly rest: SharePointRestClient;
  readonly folders: FolderService;

  private libraryPromise?: Promise<LibraryContext>;

  constructor(config: SharePointConfig, fetchImpl?: typeof fetch) {
    this.config = resolveConfig(config);
    this.rest = new SharePointRestClient({
      siteUrl: this.config.siteUrl,
      tokenProvider: this.config.tokenProvider,
      scopes: this.config.scopes,
      fetchImpl,
    });
    this.folders = new FolderService(this.rest, () => this.getLibrary());
  }

  get tokenProvider(): TokenProvider {
    return this.config.tokenProvider;
  }

  get cacheScope(): string {
    return this.config.libraryName ?? "default";
  }

  /** Tìm document library một lần rồi dùng lại cho mọi lần listChildren. */
  async getLibrary(): Promise<LibraryContext> {
    this.libraryPromise ??= resolveLibrary(
      this.rest,
      this.config.libraryName ?? "",
    );
    return this.libraryPromise;
  }
}
