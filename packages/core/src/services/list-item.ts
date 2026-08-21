import type { GraphClient } from "../graph/client";
import { itemUrl, siteResourcePath } from "../graph/paths";
import { mapDriveItem, type GraphCollection, type GraphDriveItem } from "../mappers/item";
import { isVisibleListColumn, mapListColumn, mapListItemFields } from "../mappers/list-item";
import type { ListColumn, ListItemFields, SharePointItem } from "../types/models";

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
  driveItem?: GraphDriveItem;
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

  /**
   * Official SharePoint library listing: GET /sites/{site}/lists/{list}/items?$expand=driveItem
   * Security-trimmed to the signed-in user — same source as the SharePoint list view.
   */
  async listAccessibleDriveItems(signal?: AbortSignal): Promise<SharePointItem[]> {
    const listId = await this.resolveListId(signal);
    if (!listId) return [];
    const driveId = await this.getDriveId();

    const items: SharePointItem[] = [];
    let path: string | undefined = siteResourcePath(this.siteId, `lists/${listId}/items`);
    let absoluteUrl = false;

    while (path) {
      const page: GraphCollection<GraphListItem> = await this.graph.get<GraphCollection<GraphListItem>>(path, {
        absoluteUrl,
        query: absoluteUrl
          ? undefined
          : {
              $top: 200,
              $expand:
                "driveItem($select=id,name,size,webUrl,file,folder,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,parentReference,publication,eTag,sharepointIds)",
            },
        signal,
      });

      for (const row of page.value ?? []) {
        const driveItem = row.driveItem;
        if (!driveItem?.id || !driveItem.name) continue;
        if (driveItem.parentReference?.driveId && driveItem.parentReference.driveId !== driveId) continue;
        try {
          items.push(mapDriveItem(driveItem, driveId));
        } catch {
          // Skip incomplete Graph rows.
        }
      }

      path = page["@odata.nextLink"];
      absoluteUrl = true;
    }

    return items;
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
