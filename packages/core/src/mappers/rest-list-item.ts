import type { SharePointItem } from "../types/models";
import type { RestFile, RestFolder, RestListItem } from "../types/rest";
import { mapRestItem } from "../utils/map-rest-item";
import { requireUniqueId } from "../utils/require-unique-id";

const SKIP_FIELD_KEYS = new Set(["File", "Folder"]);

function isFolderItem(item: RestListItem): boolean {
  const fs = item.FileSystemObjectType;
  const obj = item.FSObjType;
  return fs === 1 || obj === 1 || obj === "1";
}

function toSize(length: number | string | undefined): number | undefined {
  if (length === undefined || length === "") return undefined;
  const n = typeof length === "number" ? length : Number(length);
  return Number.isFinite(n) ? n : undefined;
}

function listItemFields(item: RestListItem): Record<string, unknown> {
  const fields: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(item)) {
    if (SKIP_FIELD_KEYS.has(key) || key.startsWith("@odata") || key.startsWith("odata.")) {
      continue;
    }
    fields[key] = value;
  }
  return fields;
}

/**
 * JSON list item (+ File/Folder expand) → SharePointItem.
 * Folder hệ thống Forms không đưa lên UI.
 */
export function mapRestListItem(item: RestListItem): SharePointItem | undefined {
  const folder = isFolderItem(item);
  const file = item.File as RestFile | undefined;
  const folderObj = item.Folder as RestFolder | undefined;
  const name = folder
    ? (folderObj?.Name ?? item.FileLeafRef ?? item.Title)
    : (file?.Name ?? item.FileLeafRef ?? item.Title);
  if (!name || (folder && name === "Forms")) return undefined;

  const uniqueId = folder ? folderObj?.UniqueId : file?.UniqueId;
  const id = uniqueId ?? (typeof item.GUID === "string" ? item.GUID : undefined);

  return {
    ...mapRestItem(
      folder ? "folder" : "file",
      name,
      requireUniqueId(id, folder ? "folder" : "file"),
      folder ? folderObj?.TimeLastModified : file?.TimeLastModified,
      folder ? undefined : toSize(file?.Length),
    ),
    fields: listItemFields(item),
  };
}
