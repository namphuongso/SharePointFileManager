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
  const actionName = raw.action ? Object.keys(raw.action)[0] : raw.access ? "access" : "activity";
  const itemName = raw.driveItem?.name;
  return {
    id: raw.id,
    action: actionName,
    actor,
    timestamp: raw.activityDateTime ?? raw.times?.recordedTime,
    description: itemName ? `${actionName} · ${itemName}` : actionName,
  };
}

function mapActor(identity?: GraphIdentitySet): UserInfo | undefined {
  if (!identity?.user) return undefined;
  return {
    id: identity.user.id,
    displayName: identity.user.displayName,
    email: identity.user.email,
  };
}
