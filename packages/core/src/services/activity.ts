import type { GraphClient } from "../graph/client";
import { itemUrl } from "../graph/paths";
import type { GraphCollection, GraphIdentitySet } from "../mappers/item";
import type { DriveItemActivity, UserInfo } from "../types/models";

interface GraphDriveItemActivity {
  id?: string;
  action?: Record<string, unknown>;
  access?: Record<string, unknown>;
  actor?: GraphIdentitySet;
  activityDateTime?: string;
  times?: { recordedTime?: string };
  driveItem?: { name?: string };
}

export class ActivityService {
  constructor(
    private readonly graph: GraphClient,
    private readonly getDriveId: () => Promise<string>,
  ) {}

  async list(itemId: string, signal?: AbortSignal): Promise<DriveItemActivity[]> {
    const driveId = await this.getDriveId();
    const result = await this.graph.get<GraphCollection<GraphDriveItemActivity>>(
      `${itemUrl(driveId, itemId)}/activities`,
      { signal },
    );
    return (result.value ?? [])
      .map(mapActivity)
      .filter((activity): activity is DriveItemActivity => Boolean(activity?.id));
  }
}

function mapActivity(raw: GraphDriveItemActivity): DriveItemActivity | undefined {
  if (!raw.id) return undefined;
  const actor = mapActor(raw.actor);
  const actionName = normalizeActionName(raw.action, raw.access);
  const itemName = raw.driveItem?.name;
  return {
    id: raw.id,
    action: actionName,
    actor,
    timestamp: raw.activityDateTime ?? raw.times?.recordedTime,
    description: itemName ? `${actionName} · ${itemName}` : actionName,
  };
}

function normalizeActionName(
  action?: Record<string, unknown>,
  access?: Record<string, unknown>,
): string {
  if (action) {
    const key = Object.keys(action)[0];
    if (key) {
      if (key === "edit" || key === "update") return "edited";
      if (key === "create") return "created";
      if (key === "delete") return "deleted";
      if (key === "move") return "moved";
      if (key === "rename") return "renamed";
      if (key === "share") return "shared";
      if (key === "comment") return "commented";
      return key;
    }
  }
  if (access) return "accessed";
  return "activity";
}

function mapActor(identity?: GraphIdentitySet): UserInfo | undefined {
  if (!identity?.user) return undefined;
  return {
    id: identity.user.id,
    displayName: identity.user.displayName,
    email: identity.user.email,
  };
}
