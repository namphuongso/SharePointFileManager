import type { SharePointItem, UserInfo } from "../types/models";

export interface GraphIdentitySet {
  user?: { id?: string; displayName?: string; email?: string };
  application?: { id?: string; displayName?: string };
}

export interface GraphDriveItem {
  id?: string;
  name?: string;
  size?: number;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  webUrl?: string;
  eTag?: string;
  file?: { mimeType?: string };
  folder?: { childCount?: number };
  createdBy?: GraphIdentitySet;
  lastModifiedBy?: GraphIdentitySet;
  parentReference?: { id?: string; driveId?: string; path?: string };
  "@microsoft.graph.downloadUrl"?: string;
  thumbnails?: Array<{ small?: { url?: string }; medium?: { url?: string }; large?: { url?: string } }>;
}

export interface GraphCollection<T> {
  value?: T[];
  "@odata.nextLink"?: string;
}

function mapUser(identity?: GraphIdentitySet): UserInfo | undefined {
  if (!identity?.user) return undefined;
  return {
    id: identity.user.id,
    displayName: identity.user.displayName,
    email: identity.user.email,
  };
}

export function mapDriveItem(item: GraphDriveItem, fallbackDriveId?: string): SharePointItem {
  if (!item.id || !item.name) {
    throw new Error("Graph driveItem is missing id or name");
  }

  const isFolder = Boolean(item.folder);
  return {
    id: item.id,
    name: item.name,
    type: isFolder ? "folder" : "file",
    size: item.size,
    createdDateTime: item.createdDateTime,
    lastModifiedDateTime: item.lastModifiedDateTime,
    createdBy: mapUser(item.createdBy),
    lastModifiedBy: mapUser(item.lastModifiedBy),
    webUrl: item.webUrl,
    mimeType: item.file?.mimeType,
    parentId: item.parentReference?.id,
    driveId: item.parentReference?.driveId ?? fallbackDriveId,
    eTag: item.eTag,
    childCount: item.folder?.childCount,
    downloadUrl: item["@microsoft.graph.downloadUrl"],
    thumbnailUrl:
      item.thumbnails?.[0]?.medium?.url ??
      item.thumbnails?.[0]?.small?.url ??
      item.thumbnails?.[0]?.large?.url,
    canPreview: Boolean(item.file),
  };
}
