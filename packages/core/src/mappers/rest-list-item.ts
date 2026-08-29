import type { SharePointItem } from "../types/models";
import type { RestFile, RestFolder, RestListItem } from "../types/rest";
import { mapRestItem } from "../utils/map-rest-item";
import { requireUniqueId } from "../utils/require-unique-id";

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

function toCount(value: number | string | undefined): number | undefined {
  if (value === undefined || value === "") return undefined;
  const count = typeof value === "number" ? value : Number(value);
  return Number.isFinite(count) && count >= 0 ? count : undefined;
}

/** Chỉ giữ cột được FieldService chọn; Id (OData) → ID. */
function listItemFields(
  item: RestListItem,
  fieldInternalNames: readonly string[],
): Record<string, unknown> {
  const fields: Record<string, unknown> = {};
  const allowed = new Set(fieldInternalNames);
  for (const [key, value] of Object.entries(item)) {
    if (!allowed.has(key)) continue;
    fields[key === "Id" ? "ID" : key] = value;
  }
  return fields;
}

/**
 * JSON list item (+ File/Folder expand) → SharePointItem.
 * Folder hệ thống Forms không đưa lên UI.
 */
export function mapRestListItem(
  item: RestListItem,
  fieldInternalNames: readonly string[],
): SharePointItem | undefined {
  const folder = isFolderItem(item);
  const file = item.File as RestFile | undefined;
  const folderObj = item.Folder as RestFolder | undefined;
  const name = folder
    ? (folderObj?.Name ?? item.FileLeafRef ?? item.Title)
    : (file?.Name ?? item.FileLeafRef ?? item.Title);
  if (!name || (folder && name === "Forms")) return undefined;

  const uniqueId =
    (typeof item.UniqueId === "string" ? item.UniqueId : undefined) ??
    (folder ? folderObj?.UniqueId : file?.UniqueId);

  const base = mapRestItem(
    folder ? "folder" : "file",
    name,
    requireUniqueId(uniqueId, folder ? "folder" : "file"),
    folder ? folderObj?.TimeLastModified : file?.TimeLastModified,
    folder ? undefined : toSize(file?.Length),
  );

  return {
    ...base,
    ...(folder && folderObj ? { childItemCount: toCount(folderObj.ItemCount) } : {}),
    fields: listItemFields(item, fieldInternalNames),
  };
}
