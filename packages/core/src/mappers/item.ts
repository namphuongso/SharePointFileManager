import { extractMetadataFromListItem } from "../mappers/list-item";
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
  publication?: { level?: string; versionId?: string };
  listItem?: {
    id?: string;
    contentType?: { id?: string; name?: string };
    fields?: Record<string, unknown>;
  };
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
  const listMeta = extractMetadataFromListItem(item.listItem);
  const sensitivityLabel = extractSensitivityLabel(item.listItem?.fields);
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
    listItemId: listMeta.listItemId,
    contentType: listMeta.contentType,
    metadata: listMeta.metadata,
    sensitivityLabel,
    capabilities: {
      isCheckedOut: item.publication?.level?.toLowerCase() === "checkout",
    },
  };
}

function extractSensitivityLabel(fields?: Record<string, unknown>): string | undefined {
  if (!fields) return undefined;
  const candidates = [
    fields._ComplianceTag,
    fields.Sensitivity,
    fields.SensitivityLabel,
    fields._SensitivityLabel,
  ];
  for (const candidate of candidates) {
    if (typeof candidate === "string" && candidate.trim()) return candidate.trim();
    if (candidate && typeof candidate === "object") {
      const record = candidate as Record<string, unknown>;
      if (typeof record.Label === "string" && record.Label.trim()) return record.Label.trim();
      if (typeof record.displayName === "string" && record.displayName.trim()) return record.displayName.trim();
    }
  }
  return undefined;
}
