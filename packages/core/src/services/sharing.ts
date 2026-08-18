import type { GraphClient } from "../graph/client";
import { itemUrl } from "../graph/paths";
import { mapPermission } from "../mappers/permission";
import type {
  CreateLinkOptions,
  InviteOptions,
  SharePointPermission,
} from "../types/models";

export class SharingService {
  constructor(
    private readonly graph: GraphClient,
    private readonly getDriveId: () => Promise<string>,
  ) {}

  async invite(itemId: string, options: InviteOptions, signal?: AbortSignal): Promise<SharePointPermission[]> {
    const driveId = await this.getDriveId();
    const result = await this.graph.post<{ value?: unknown[] } | unknown[]>(
      `${itemUrl(driveId, itemId)}/invite`,
      {
        recipients: options.recipients.map((recipient) =>
          recipient.objectId ? { objectId: recipient.objectId } : { email: recipient.email },
        ),
        roles: [options.role],
        message: options.message,
        requireSignIn: options.requireSignIn ?? true,
        sendInvitation: options.sendInvitation ?? true,
      },
      { signal },
    );

    const value = Array.isArray(result) ? result : (result.value ?? []);
    return value.map((permission) => mapPermission(permission as Parameters<typeof mapPermission>[0]));
  }

  async createLink(
    itemId: string,
    options: CreateLinkOptions,
    signal?: AbortSignal,
  ): Promise<SharePointPermission> {
    const driveId = await this.getDriveId();
    const permission = await this.graph.post(
      `${itemUrl(driveId, itemId)}/createLink`,
      {
        type: options.type,
        scope: options.scope,
        expirationDateTime: options.expirationDateTime,
      },
      { signal },
    );
    return mapPermission(permission as Parameters<typeof mapPermission>[0]);
  }
}

export class PermissionService {
  constructor(
    private readonly graph: GraphClient,
    private readonly getDriveId: () => Promise<string>,
  ) {}

  async list(itemId: string, signal?: AbortSignal): Promise<SharePointPermission[]> {
    const driveId = await this.getDriveId();
    const result = await this.graph.get<{ value?: unknown[] }>(`${itemUrl(driveId, itemId)}/permissions`, {
      signal,
    });
    return (result.value ?? []).map((permission) =>
      mapPermission(permission as Parameters<typeof mapPermission>[0]),
    );
  }

  async update(
    itemId: string,
    permissionId: string,
    roles: string[],
    signal?: AbortSignal,
  ): Promise<SharePointPermission> {
    const driveId = await this.getDriveId();
    const permission = await this.graph.patch(
      `${itemUrl(driveId, itemId)}/permissions/${permissionId}`,
      { roles },
      { signal },
    );
    return mapPermission(permission as Parameters<typeof mapPermission>[0]);
  }

  async remove(itemId: string, permissionId: string, signal?: AbortSignal): Promise<void> {
    const driveId = await this.getDriveId();
    await this.graph.delete(`${itemUrl(driveId, itemId)}/permissions/${permissionId}`, { signal });
  }

}
