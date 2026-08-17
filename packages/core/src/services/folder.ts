import { SharePointError, SharePointErrorCode } from "../errors/sharepoint-error";
import type { GraphClient } from "../graph/client";
import { childrenUrl, itemUrl } from "../graph/paths";
import { mapDriveItem, type GraphCollection, type GraphDriveItem } from "../mappers/item";
import type { CopyMoveOptions, PagedResult, SharePointItem } from "../types/models";

export class FolderService {
  constructor(
    private readonly graph: GraphClient,
    private readonly getDriveId: () => Promise<string>,
  ) {}

  async get(itemId: string, signal?: AbortSignal): Promise<SharePointItem> {
    const driveId = await this.getDriveId();
    const item = await this.graph.get<GraphDriveItem>(itemUrl(driveId, itemId), { signal });
    return mapDriveItem(item, driveId);
  }

  async listChildren(
    folderId: string,
    options: { top?: number; nextLink?: string; signal?: AbortSignal } = {},
  ): Promise<PagedResult<SharePointItem>> {
    const driveId = await this.getDriveId();
    const items: SharePointItem[] = [];
    let path: string | undefined = options.nextLink ?? childrenUrl(driveId, folderId);
    let absoluteUrl = Boolean(options.nextLink);

    while (path) {
      const result: GraphCollection<GraphDriveItem> = await this.graph.get<GraphCollection<GraphDriveItem>>(path, {
        absoluteUrl,
        query: absoluteUrl ? undefined : { $top: options.top ?? 200 },
        signal: options.signal,
      });

      for (const raw of result.value ?? []) {
        try {
          items.push(mapDriveItem(raw, driveId));
        } catch {
          // Skip incomplete Graph rows instead of dropping the whole folder.
        }
      }

      path = options.nextLink ? undefined : result["@odata.nextLink"];
      absoluteUrl = true;
    }

    items.sort((a, b) => {
      if (a.type !== b.type) return a.type === "folder" ? -1 : 1;
      return a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
    });

    return { items };
  }

  async create(parentId: string, name: string, signal?: AbortSignal): Promise<SharePointItem> {
    const driveId = await this.getDriveId();
    const item = await this.graph.post<GraphDriveItem>(
      childrenUrl(driveId, parentId),
      {
        name,
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail",
      },
      { signal },
    );
    return mapDriveItem(item, driveId);
  }

  async rename(itemId: string, name: string, signal?: AbortSignal): Promise<SharePointItem> {
    const driveId = await this.getDriveId();
    const item = await this.graph.patch<GraphDriveItem>(itemUrl(driveId, itemId), { name }, { signal });
    return mapDriveItem(item, driveId);
  }

  async delete(itemId: string, signal?: AbortSignal): Promise<void> {
    const driveId = await this.getDriveId();
    await this.graph.delete(itemUrl(driveId, itemId), { signal });
  }

  async copy(options: CopyMoveOptions): Promise<void> {
    const driveId = await this.getDriveId();
    const response = await this.graph.request<Response>({
      path: `${itemUrl(driveId, options.itemId)}/copy`,
      method: "POST",
      body: {
        parentReference: { id: options.destinationParentId },
        name: options.newName,
      },
      raw: true,
      signal: options.signal,
    });
    await waitForAsyncOperation(this.graph, response, options.signal);
  }

  async move(options: CopyMoveOptions): Promise<SharePointItem> {
    const driveId = await this.getDriveId();
    const body: Record<string, unknown> = {
      parentReference: { id: options.destinationParentId },
    };
    if (options.newName) body.name = options.newName;
    const item = await this.graph.patch<GraphDriveItem>(itemUrl(driveId, options.itemId), body, {
      signal: options.signal,
    });
    return mapDriveItem(item, driveId);
  }
}

async function waitForAsyncOperation(
  graph: GraphClient,
  response: Response,
  signal?: AbortSignal,
): Promise<void> {
  const monitorUrl = response.headers.get("Location");
  if (!monitorUrl) return;

  for (let attempt = 0; attempt < 40; attempt += 1) {
    if (signal?.aborted) {
      throw new SharePointError({
        code: SharePointErrorCode.Cancelled,
        message: "Copy was cancelled",
      });
    }

    const status = await graph.get<{ status?: string; error?: { message?: string } }>(monitorUrl, {
      absoluteUrl: true,
      signal,
    });

    if (!status.status || status.status === "completed" || status.status === "succeeded") {
      return;
    }
    if (status.status === "failed") {
      throw new SharePointError({
        code: SharePointErrorCode.Unknown,
        message: status.error?.message ?? "Copy operation failed",
      });
    }

    await new Promise((resolve) => setTimeout(resolve, 400));
  }
}
