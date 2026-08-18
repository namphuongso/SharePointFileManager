import type { GraphClient } from "../graph/client";
import { itemUrl } from "../graph/paths";
import { mapDriveItem, type GraphCollection, type GraphDriveItem } from "../mappers/item";
import type { PagedResult, SearchFilters, SearchOptions, SearchScope, SharePointItem } from "../types/models";

interface GraphSearchHit {
  resource?: GraphDriveItem & { "@odata.type"?: string };
}

interface GraphSearchResponse {
  value?: Array<{
    hitsContainers?: Array<{
      hits?: GraphSearchHit[];
      moreResultsAvailable?: boolean;
    }>;
  }>;
}

export class SearchService {
  private driveWebUrlPromise?: Promise<string | undefined>;

  constructor(
    private readonly graph: GraphClient,
    private readonly getDriveId: () => Promise<string>,
    private readonly rootItemId: string,
  ) {}

  async search(options: SearchOptions): Promise<PagedResult<SharePointItem>> {
    const scope: SearchScope = options.scope ?? "folder";
    if (scope === "library") {
      return this.searchLibrary(options);
    }
    return this.searchFolder(options);
  }

  private async searchFolder(options: SearchOptions): Promise<PagedResult<SharePointItem>> {
    if (options.nextLink) {
      return this.searchFolderByUrl(options.nextLink, options.signal);
    }

    const driveId = await this.getDriveId();
    const folderId = options.folderId ?? this.rootItemId;
    const encoded = options.query.replace(/'/g, "''");
    const result = await this.graph.get<GraphCollection<GraphDriveItem>>(
      `${itemUrl(driveId, folderId)}/search(q='${encoded}')`,
      {
        query: { $top: options.top ?? 25 },
        signal: options.signal,
      },
    );

    return {
      items: (result.value ?? []).map((item) => mapDriveItem(item, driveId)),
      nextLink: result["@odata.nextLink"],
    };
  }

  private async searchFolderByUrl(nextLink: string, signal?: AbortSignal): Promise<PagedResult<SharePointItem>> {
    const driveId = await this.getDriveId();
    const result = await this.graph.get<GraphCollection<GraphDriveItem>>(nextLink, {
      absoluteUrl: true,
      signal,
    });
    return {
      items: (result.value ?? []).map((item) => mapDriveItem(item, driveId)),
      nextLink: result["@odata.nextLink"],
    };
  }

  private async searchLibrary(options: SearchOptions): Promise<PagedResult<SharePointItem>> {
    const driveId = await this.getDriveId();
    const pageSize = options.top ?? 25;
    const from = options.from ?? 0;
    const pathPrefix = await this.resolvePathPrefix(options.folderId, options.signal);
    const queryString = buildSearchKql(options.query, pathPrefix, options.filters);

    const result = await this.graph.post<GraphSearchResponse>(
      "/search/query",
      {
        requests: [
          {
            entityTypes: ["driveItem"],
            query: { queryString: queryString || "*" },
            from,
            size: pageSize,
            fields: [
              "id",
              "name",
              "webUrl",
              "size",
              "lastModifiedDateTime",
              "createdDateTime",
              "parentReference",
              "file",
              "folder",
            ],
          },
        ],
      },
      { signal: options.signal },
    );

    const hitsContainer = result.value?.[0]?.hitsContainers?.[0];
    const hits = hitsContainer?.hits ?? [];
    const items = hits
      .map((hit) => hit.resource)
      .filter((resource): resource is GraphDriveItem => Boolean(resource?.id && resource?.name))
      .map((item) => mapDriveItem(item, driveId));

    const moreResultsAvailable = hitsContainer?.moreResultsAvailable === true;
    return {
      items,
      nextLink: moreResultsAvailable ? encodeLibraryPage(from + pageSize) : undefined,
    };
  }

  private async resolvePathPrefix(folderId: string | undefined, signal?: AbortSignal): Promise<string | undefined> {
    const driveId = await this.getDriveId();
    const targetId = folderId && folderId !== this.rootItemId ? folderId : undefined;
    if (targetId) {
      const item = await this.graph.get<GraphDriveItem>(itemUrl(driveId, targetId), {
        query: { $select: "webUrl" },
        signal,
      });
      return item.webUrl;
    }
    return this.getDriveWebUrl(signal);
  }

  private async getDriveWebUrl(signal?: AbortSignal): Promise<string | undefined> {
    this.driveWebUrlPromise ??= this.loadDriveWebUrl(signal);
    return this.driveWebUrlPromise;
  }

  private async loadDriveWebUrl(signal?: AbortSignal): Promise<string | undefined> {
    try {
      const driveId = await this.getDriveId();
      const root = await this.graph.get<GraphDriveItem>(`/drives/${driveId}/root`, {
        query: { $select: "webUrl" },
        signal,
      });
      return root.webUrl;
    } catch {
      return undefined;
    }
  }
}

export function buildSearchKql(query: string, pathPrefix?: string, filters?: SearchFilters): string {
  const parts: string[] = [];
  const trimmed = query.trim();
  if (trimmed) parts.push(trimmed);
  if (pathPrefix) parts.push(`path:"${escapeKql(pathPrefix)}"`);
  if (filters?.fileType?.trim()) parts.push(`filetype:${filters.fileType.trim()}`);
  if (filters?.modifiedAfter) parts.push(`LastModifiedTime>=${filters.modifiedAfter}`);
  if (filters?.modifiedBefore) parts.push(`LastModifiedTime<=${filters.modifiedBefore}`);
  if (filters?.author?.trim()) parts.push(`author:"${escapeKql(filters.author.trim())}"`);
  return parts.join(" AND ");
}

export function decodeLibraryPage(nextLink?: string): number {
  if (!nextLink?.startsWith("library-search:")) return 0;
  const value = Number(nextLink.slice("library-search:".length));
  return Number.isFinite(value) ? value : 0;
}

function encodeLibraryPage(from: number): string {
  return `library-search:${from}`;
}

function escapeKql(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}
