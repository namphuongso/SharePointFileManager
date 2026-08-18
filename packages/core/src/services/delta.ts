import type { GraphClient } from "../graph/client";
import type { GraphCollection, GraphDriveItem } from "../mappers/item";
import { isSharePointError } from "../errors/sharepoint-error";

export interface DeltaPage<T> {
  items: T[];
  nextLink?: string;
  deltaLink?: string;
  hasMore: boolean;
}

export interface DeltaChange<T> {
  item: T;
  deleted: boolean;
}

export class DeltaService {
  constructor(private readonly graph: GraphClient) {}

  async start(driveId: string, itemId = "root", signal?: AbortSignal): Promise<DeltaPage<GraphDriveItem>> {
    if (itemId !== "root") {
      throw new Error("Microsoft Graph delta is supported for the drive root only");
    }
    return this.read(`/drives/${driveId}/root/delta`, signal);
  }

  async next<T = GraphDriveItem>(nextLink: string, signal?: AbortSignal): Promise<DeltaPage<T>> {
    return this.read<T>(nextLink, signal);
  }

  async sync(
    driveId: string,
    itemId = "root",
    options: { deltaLink?: string; signal?: AbortSignal; maxPages?: number } = {},
  ): Promise<{ changes: DeltaChange<GraphDriveItem>[]; deltaLink?: string }> {
    let page: DeltaPage<GraphDriveItem>;
    try {
      page = options.deltaLink
        ? await this.read(options.deltaLink, options.signal)
        : await this.start(driveId, itemId, options.signal);
    } catch (error) {
      if (!options.deltaLink || !isSharePointError(error) || error.status !== 410) throw error;
      // Graph invalidated the saved token. Restart from the root snapshot as
      // documented instead of retrying a permanently invalid delta URL.
      page = await this.start(driveId, itemId, options.signal);
    }
    const changes: DeltaChange<GraphDriveItem>[] = [];
    let pages = 1;
    while (true) {
      changes.push(...page.items.map((item) => ({ item, deleted: Boolean((item as GraphDriveItem & { deleted?: unknown }).deleted) })));
      if (!page.nextLink || (options.maxPages !== undefined && pages >= options.maxPages)) break;
      page = await this.next(page.nextLink, options.signal);
      pages += 1;
    }
    return { changes, deltaLink: page.deltaLink };
  }

  private async read<T>(url: string, signal?: AbortSignal): Promise<DeltaPage<T>> {
    const result = await this.graph.get<GraphCollection<T> & { "@odata.deltaLink"?: string }>(url, {
      absoluteUrl: /^https?:\/\//i.test(url), signal,
    });
    return {
      items: result.value ?? [],
      nextLink: result["@odata.nextLink"],
      deltaLink: result["@odata.deltaLink"],
      hasMore: Boolean(result["@odata.nextLink"]),
    };
  }
}
