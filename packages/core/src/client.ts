import type { TokenProvider } from "./auth/token-provider";
import { resolveConfig } from "./config/resolve-config";
import { SharePointError, SharePointErrorCode } from "./errors/sharepoint-error";
import { GraphClient } from "./graph/client";
import { DriveService, findDriveByName } from "./services/drive";
import { SiteService } from "./services/site";
import { FileService } from "./services/file";
import { FolderService } from "./services/folder";
import { PeopleService } from "./services/people";
import { PermissionService, SharingService } from "./services/sharing";
import { SearchService } from "./services/search";
import { CheckoutService } from "./services/checkout";
import { UploadService } from "./services/upload";
import { ListItemService } from "./services/list-item";
import { ActivityService } from "./services/activity";
import { DeltaService } from "./services/delta";
import type { OfficeFileKind, ResolvedSharePointConfig, SharePointConfig, SharePointItem } from "./types/models";

export class SharePointClient {
  readonly config: ResolvedSharePointConfig;
  readonly graph: GraphClient;
  readonly sites: SiteService;
  readonly drives: DriveService;
  readonly files: FileService;
  readonly folders: FolderService;
  readonly upload: UploadService;
  readonly sharing: SharingService;
  readonly permissions: PermissionService;
  readonly search: SearchService;
  readonly people: PeopleService;
  readonly checkout: CheckoutService;
  readonly listItems: ListItemService;
  readonly activities: ActivityService;
  readonly delta: DeltaService;
  /** Spec-aligned aliases retained alongside the original service names. */
  readonly driveItems: FolderService;
  readonly lists: ListItemService;

  private driveIdPromise?: Promise<string>;

  constructor(config: SharePointConfig, fetchImpl?: typeof fetch) {
    this.config = resolveConfig(config);
    this.graph = new GraphClient({
      baseUrl: this.config.graphBaseUrl,
      tokenProvider: this.config.tokenProvider,
      scopes: this.config.scopes,
      fetchImpl,
    });
    this.sites = new SiteService(this.graph);
    this.drives = new DriveService(this.graph, this.config.siteId);
    this.folders = new FolderService(this.graph, () => this.getDriveId());
    this.files = new FileService(this.graph, () => this.getDriveId(), this.folders);
    this.upload = new UploadService(this.graph, () => this.getDriveId());
    this.sharing = new SharingService(this.graph, () => this.getDriveId());
    this.permissions = new PermissionService(this.graph, () => this.getDriveId());
    this.search = new SearchService(this.graph, () => this.getDriveId(), this.config.rootItemId);
    this.people = new PeopleService(this.graph);
    this.checkout = new CheckoutService(this.graph, () => this.getDriveId(), this.folders);
    this.listItems = new ListItemService(
      this.graph,
      this.config.siteId,
      () => this.getDriveId(),
      this.config.listId,
    );
    this.activities = new ActivityService(this.graph, () => this.getDriveId());
    this.delta = new DeltaService(this.graph);
    this.driveItems = this.folders;
    this.lists = this.listItems;
  }

  get tokenProvider(): TokenProvider {
    return this.config.tokenProvider;
  }

  /** Cache key that distinguishes libraries on the same site. */
  get cacheScope(): string {
    return this.config.driveId ?? this.config.listId ?? this.config.libraryName ?? "default";
  }

  async getDriveId(): Promise<string> {
    if (this.config.driveId) return this.config.driveId;
    this.driveIdPromise ??= this.resolveDriveId();
    return this.driveIdPromise;
  }

  private async resolveDriveId(): Promise<string> {
    if (this.config.listId) {
      return (await this.drives.getDriveByList(this.config.listId)).id;
    }

    if (this.config.libraryName) {
      const drives = await this.drives.listDrives();
      const match = findDriveByName(drives, this.config.libraryName);
      if (!match) {
        throw new SharePointError({
          code: SharePointErrorCode.NotFound,
          message: `SharePoint library "${this.config.libraryName}" was not found on this site`,
        });
      }
      return match.id;
    }

    return (await this.drives.getDefaultDrive()).id;
  }

  /** Create an empty Office file placeholder (SharePoint opens it in Office Online). */
  async createOfficeFile(
    parentId: string,
    kind: OfficeFileKind,
    signal?: AbortSignal,
  ): Promise<SharePointItem> {
    const names: Record<OfficeFileKind, string> = {
      word: "Document.docx",
      excel: "Workbook.xlsx",
      powerpoint: "Presentation.pptx",
    };
    return this.upload.createPlaceholder(parentId, names[kind], signal);
  }
}
