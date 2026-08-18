import type { GraphClient } from "../graph/client";
import { itemUrl } from "../graph/paths";
import type { SharePointItem } from "../types/models";
import type { FolderService } from "./folder";

export class CheckoutService {
  constructor(
    private readonly graph: GraphClient,
    private readonly getDriveId: () => Promise<string>,
    private readonly folders: FolderService,
  ) {}

  async checkout(itemId: string, signal?: AbortSignal): Promise<SharePointItem> {
    const driveId = await this.getDriveId();
    await this.graph.post(`${itemUrl(driveId, itemId)}/checkout`, undefined, { signal });
    return this.enrichCheckout(
      await this.folders.get(itemId, { signal, expandListItem: true }),
      signal,
    );
  }

  async checkin(itemId: string, comment?: string, signal?: AbortSignal): Promise<SharePointItem> {
    const driveId = await this.getDriveId();
    await this.graph.post(
      `${itemUrl(driveId, itemId)}/checkin`,
      { comment: comment ?? "" },
      { signal },
    );
    return this.enrichCheckout(
      await this.folders.get(itemId, { signal, expandListItem: true }),
      signal,
    );
  }

  async discardCheckout(itemId: string, signal?: AbortSignal): Promise<SharePointItem> {
    const driveId = await this.getDriveId();
    await this.graph.post(`${itemUrl(driveId, itemId)}/discardCheckout`, undefined, { signal });
    return this.enrichCheckout(
      await this.folders.get(itemId, { signal, expandListItem: true }),
      signal,
    );
  }

  async enrichCheckout(item: SharePointItem, signal?: AbortSignal): Promise<SharePointItem> {
    if (item.type !== "file") return item;
    const driveId = await this.getDriveId();
    try {
      const fields = await this.graph.get<Record<string, unknown>>(
        `${itemUrl(driveId, item.id)}/listItem/fields`,
        {
          query: { $select: "CheckoutUserId,CheckoutUserLookupId,CheckedOutUserId" },
          signal,
        },
      );
      const checkoutUserId = text(fields.CheckoutUserId) || text(fields.CheckedOutUserId);
      const isCheckedOut = Boolean(checkoutUserId) || item.capabilities?.isCheckedOut;
      return {
        ...item,
        capabilities: {
          ...item.capabilities,
          isCheckedOut,
          checkedOutBy: checkoutUserId ? { id: checkoutUserId } : item.capabilities?.checkedOutBy,
          checkedOutByMe: item.capabilities?.checkedOutByMe,
        },
      };
    } catch {
      return item;
    }
  }
}

function text(value: unknown): string {
  return typeof value === "string" ? value.trim() : typeof value === "number" ? String(value) : "";
}
