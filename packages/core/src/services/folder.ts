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

  async get(
    itemId: string,
    options: { signal?: AbortSignal; expandListItem?: boolean } = {},
  ): Promise<SharePointItem> {
    const driveId = await this.getDriveId();
    const expand = buildExpandQuery(options.expandListItem);
    const item = await this.graph.get<GraphDriveItem>(itemUrl(driveId, itemId), {
      query: {
        $select:
          "id,name,size,webUrl,file,folder,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,parentReference,publication,eTag,sharepointIds",
        ...(expand ? { $expand: expand } : {}),
      },
      signal: options.signal,
    });
    return mapDriveItem(item, driveId);
  }

  async listChildren(
    folderId: string,
    options: { top?: number; nextLink?: string; signal?: AbortSignal; expandListItem?: boolean } = {},
  ): Promise<PagedResult<SharePointItem>> {
    if (!options.nextLink && options.top !== undefined) {
      return this.listChildrenPage(folderId, options);
    }

    const driveId = await this.getDriveId();
    const items: SharePointItem[] = [];
    let path: string | undefined = options.nextLink ?? childrenUrl(driveId, folderId);
    let absoluteUrl = Boolean(options.nextLink);

    while (path) {
      const result: GraphCollection<GraphDriveItem> = await this.graph.get<GraphCollection<GraphDriveItem>>(path, {
        absoluteUrl,
        query: absoluteUrl
          ? undefined
          : {
              $top: options.top ?? 200,
              $select:
                "id,name,size,webUrl,file,folder,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,parentReference,publication,eTag,sharepointIds",
              ...(buildExpandQuery(options.expandListItem)
                ? { $expand: buildExpandQuery(options.expandListItem) }
                : {}),
            },
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

  async listChildrenPage(
    folderId: string,
    options: { top?: number; nextLink?: string; signal?: AbortSignal; expandListItem?: boolean } = {},
  ): Promise<PagedResult<SharePointItem>> {
    const driveId = await this.getDriveId();
    const path = options.nextLink ?? childrenUrl(driveId, folderId);
    const result = await this.graph.get<GraphCollection<GraphDriveItem>>(path, {
      absoluteUrl: Boolean(options.nextLink),
      query: options.nextLink
        ? undefined
        : {
            $top: options.top ?? 50,
            $select:
              "id,name,size,webUrl,file,folder,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,parentReference,publication,eTag,sharepointIds",
            ...(buildExpandQuery(options.expandListItem)
              ? { $expand: buildExpandQuery(options.expandListItem) }
              : {}),
          },
      signal: options.signal,
    });

    const items: SharePointItem[] = [];
    for (const raw of result.value ?? []) {
      try {
        items.push(mapDriveItem(raw, driveId));
      } catch {
        // Skip incomplete rows.
      }
    }

    items.sort((a, b) => {
      if (a.type !== b.type) return a.type === "folder" ? -1 : 1;
      return a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
    });

    return {
      items,
      nextLink: result["@odata.nextLink"],
    };
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
        parentReference: { driveId, id: options.destinationParentId },
        name: options.newName,
      },
      raw: true,
      signal: options.signal,
    });
    await waitForAsyncOperation(this.graph, response, options.signal, options.onCopyProgress);
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

function buildExpandQuery(expandListItem?: boolean): string {
  const parts = ["thumbnails($select=small,medium,large)"];
  if (expandListItem) {
    parts.push("listItem($select=id,contentType;$expand=fields)");
  }
  return parts.join(",");
}

async function waitForAsyncOperation(
  graph: GraphClient,
  response: Response,
  signal?: AbortSignal,
  onProgress?: (progress: import("../types/models").CopyOperationProgress) => void,
): Promise<void> {
  const monitorUrl = response.headers.get("Location");
  if (!monitorUrl) {
    throw new SharePointError({
      code: SharePointErrorCode.NetworkError,
      message: "Copy operation did not return the Microsoft Graph monitor URL",
    });
  }

  onProgress?.({ phase: "starting", percent: 0 });

  const deadline = Date.now() + 5 * 60_000;
  let attempt = 0;
  while (Date.now() < deadline) {
    if (signal?.aborted) {
      throw new SharePointError({
        code: SharePointErrorCode.Cancelled,
        message: "Copy was cancelled",
      });
    }

    const status = await graph.get<{
      status?: string;
      percentComplete?: number;
      percentageComplete?: number;
      error?: { message?: string };
    }>(monitorUrl, {
      absoluteUrl: true,
      signal,
    });

    const normalizedStatus = status.status?.toLowerCase();
    if (normalizedStatus === "completed" || normalizedStatus === "succeeded") {
      onProgress?.({ phase: "completed", percent: 100 });
      return;
    }
    if (normalizedStatus === "failed") {
      onProgress?.({ phase: "failed", percent: 0 });
      throw new SharePointError({
        code: SharePointErrorCode.Unknown,
        message: status.error?.message ?? "Copy operation failed",
      });
    }

    onProgress?.({
      phase: "monitoring",
      percent:
        status.percentComplete ??
        status.percentageComplete ??
        Math.min(95, Math.max(1, Math.round((attempt / 300) * 100))),
    });
    attempt += 1;
    await waitForCopyPoll(1000, signal);
  }
  onProgress?.({ phase: "failed", percent: 0 });
  throw new SharePointError({
    code: SharePointErrorCode.NetworkError,
    message: "Copy operation did not complete before the monitoring timeout",
  });
}

function waitForCopyPoll(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    let timer: ReturnType<typeof setTimeout> | undefined;
    const onAbort = () => {
      if (timer) clearTimeout(timer);
      signal?.removeEventListener("abort", onAbort);
      reject(new SharePointError({ code: SharePointErrorCode.Cancelled, message: "Copy was cancelled" }));
    };
    timer = setTimeout(() => {
      signal?.removeEventListener("abort", onAbort);
      resolve();
    }, ms);
    signal?.addEventListener("abort", onAbort, { once: true });
  });
}
