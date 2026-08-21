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
    return this.folders.get(itemId, { signal, expandListItem: true });
  }

  async checkin(itemId: string, comment?: string, signal?: AbortSignal): Promise<SharePointItem> {
    const driveId = await this.getDriveId();
    await this.graph.post(
      `${itemUrl(driveId, itemId)}/checkin`,
      { comment: comment ?? "" },
      { signal },
    );
    return this.folders.get(itemId, { signal, expandListItem: true });
  }

  async discardCheckout(itemId: string, signal?: AbortSignal): Promise<SharePointItem> {
    const driveId = await this.getDriveId();
    await this.graph.post(`${itemUrl(driveId, itemId)}/discardCheckout`, undefined, { signal });
    return this.folders.get(itemId, { signal, expandListItem: true });
  }
}
