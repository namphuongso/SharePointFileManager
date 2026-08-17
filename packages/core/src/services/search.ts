import type { GraphClient } from "../graph/client";
import { itemUrl } from "../graph/paths";
import { mapDriveItem, type GraphCollection, type GraphDriveItem } from "../mappers/item";
import type { PagedResult, SearchOptions, SharePointItem } from "../types/models";

export class SearchService {
  constructor(
    private readonly graph: GraphClient,
    private readonly getDriveId: () => Promise<string>,
    private readonly rootItemId: string,
  ) {}

  async search(options: SearchOptions): Promise<PagedResult<SharePointItem>> {
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
}
