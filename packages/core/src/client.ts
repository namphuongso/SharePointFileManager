import type { TokenProvider } from "./auth/token-provider";
import { resolveConfig } from "./config/resolve-config";
import { SharePointRestClient } from "./rest/client";
import { FolderService } from "./services/folder";
import { resolveLibrary, type LibraryContext } from "./services/library";
import type { ResolvedSharePointConfig, SharePointConfig } from "./types/models";

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
    return this.config.listId ?? this.config.libraryName ?? "default";
  }

  async getLibrary(): Promise<LibraryContext> {
    this.libraryPromise ??= resolveLibrary(this.rest, {
      libraryName: this.config.libraryName,
      listId: this.config.listId,
    });
    return this.libraryPromise;
  }
}
