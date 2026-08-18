import type { GraphClient } from "../graph/client";
import { itemUrl, siteResourcePath } from "../graph/paths";
import type { GraphCollection } from "../mappers/item";
import { isVisibleListColumn, mapListColumn, mapListItemFields } from "../mappers/list-item";
import type { ListColumn, ListItemFields } from "../types/models";

interface GraphListColumn {
  id?: string;
  name?: string;
  displayName?: string;
  hidden?: boolean;
  readOnly?: boolean;
  text?: unknown;
  choice?: unknown;
  dateTime?: unknown;
  number?: unknown;
  boolean?: unknown;
  lookup?: unknown;
  personOrGroup?: unknown;
}

interface GraphListItem {
  id?: string;
  contentType?: { id?: string; name?: string };
  fields?: Record<string, unknown>;
}

interface GraphDriveList {
  id?: string;
}

export class ListItemService {
  private listIdPromise?: Promise<string | undefined>;

  constructor(
    private readonly graph: GraphClient,
    private readonly siteId: string,
    private readonly getDriveId: () => Promise<string>,
    private readonly configuredListId?: string,
  ) {}

  async listColumns(signal?: AbortSignal): Promise<ListColumn[]> {
    const listId = await this.resolveListId(signal);
    if (!listId) return [];

    const result = await this.graph.get<GraphCollection<GraphListColumn>>(
      siteResourcePath(this.siteId, `lists/${listId}/columns`),
      {
        query: {
          $select: "id,name,displayName,hidden,readOnly,text,choice,dateTime,number,boolean,lookup,personOrGroup",
          $top: 200,
        },
        signal,
      },
    );

    return (result.value ?? [])
      .map(mapListColumn)
      .filter((column): column is ListColumn => Boolean(column))
      .filter(isVisibleListColumn);
  }

  async getFields(itemId: string, signal?: AbortSignal): Promise<ListItemFields> {
    const driveId = await this.getDriveId();
    const listItem = await this.graph.get<GraphListItem>(`${itemUrl(driveId, itemId)}/listItem`, {
      query: { $expand: "fields" },
      signal,
    });
    return mapListItemFields(itemId, listItem);
  }

  async updateFields(
    itemId: string,
    fields: Record<string, string | number | boolean | null>,
    signal?: AbortSignal,
  ): Promise<ListItemFields> {
    const driveId = await this.getDriveId();
    await this.graph.patch(`${itemUrl(driveId, itemId)}/listItem/fields`, fields, { signal });
    return this.getFields(itemId, signal);
  }

  private async resolveListId(signal?: AbortSignal): Promise<string | undefined> {
    if (this.configuredListId) return this.configuredListId;
    this.listIdPromise ??= this.loadListIdFromDrive(signal);
    return this.listIdPromise;
  }

  private async loadListIdFromDrive(signal?: AbortSignal): Promise<string | undefined> {
    try {
      const driveId = await this.getDriveId();
      const list = await this.graph.get<GraphDriveList>(`/drives/${driveId}/list`, { signal });
      return list.id;
    } catch {
      return undefined;
    }
  }
}
