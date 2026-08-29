import type { SharePointItem } from "../types/models";
import type { RestSearchCell, RestSearchRow } from "../types/search";
import { mapRestItem } from "../utils/map-rest-item";
import { requireUniqueId } from "../utils/require-unique-id";
import { searchManagedProperty } from "../services/search/search-sort";

/** Cells → map Key/Value. */
function cellsToMap(cells: RestSearchCell[] | undefined): Map<string, string> {
  const map = new Map<string, string>();
  if (!cells) return map;
  for (const cell of cells) {
    if (!cell.Key || cell.Value == null || cell.Value === "") continue;
    map.set(cell.Key, String(cell.Value));
  }
  return map;
}

function isTruthy(value: string | undefined): boolean {
  if (!value) return false;
  const v = value.trim().toLowerCase();
  return v === "true" || v === "1" || v === "yes";
}

function fileNameFromPath(path: string): string | undefined {
  const cleaned = path.replace(/\/+$/, "");
  const slash = cleaned.lastIndexOf("/");
  if (slash < 0) return cleaned || undefined;
  return cleaned.slice(slash + 1) || undefined;
}

/** Person crawl thường là chuỗi tên — gắn shape Title để UI pill dùng chung. */
function personField(name: string | undefined): { Title: string } | undefined {
  const title = name?.trim();
  return title ? { Title: title } : undefined;
}

/**
 * Một dòng Search → SharePointItem (+ fields cột option khi có managed property).
 * UniqueId bắt buộc; Forms / thư viện gốc bỏ.
 */
export function mapSearchRow(
  row: RestSearchRow,
  libraryAbsoluteUrl: string,
  fieldInternalNames: readonly string[] = [],
): SharePointItem | undefined {
  const cells = cellsToMap(row.Cells);
  const path = cells.get("Path");
  if (!path) return undefined;

  const library = libraryAbsoluteUrl.replace(/\/+$/, "").toLowerCase();
  const pathNorm = path.replace(/\/+$/, "").toLowerCase();
  if (pathNorm === library) return undefined;
  if (pathNorm.includes("/forms")) return undefined;

  const uniqueId = cells.get("UniqueId")?.replace(/[{}]/g, "");
  if (!uniqueId) return undefined;

  const folder = isTruthy(cells.get("IsContainer"));
  const name =
    cells.get("Filename") ||
    cells.get("Title") ||
    fileNameFromPath(path);
  if (!name) return undefined;

  const sizeRaw = cells.get("Size");
  const sizeNum = sizeRaw !== undefined ? Number(sizeRaw) : undefined;
  const size =
    !folder && sizeNum !== undefined && Number.isFinite(sizeNum) && sizeNum >= 0
      ? sizeNum
      : undefined;

  const childRaw = cells.get("FolderChildCount");
  const childNum = childRaw !== undefined ? Number(childRaw) : undefined;
  const childItemCount =
    folder && childNum !== undefined && Number.isFinite(childNum) && childNum >= 0
      ? childNum
      : undefined;

  const fields: Record<string, unknown> = {};
  for (const internalName of fieldInternalNames) {
    const managed = searchManagedProperty(internalName);
    const raw = cells.get(managed);
    if (raw == null || raw === "") continue;
    if (internalName === "Author" || internalName === "Editor") {
      const person = personField(raw);
      if (person) fields[internalName] = person;
      continue;
    }
    if (internalName === "ItemChildCount") {
      const n = Number(raw);
      if (Number.isFinite(n)) fields[internalName] = n;
      continue;
    }
    fields[internalName] = raw;
  }

  const base = mapRestItem(
    folder ? "folder" : "file",
    name,
    requireUniqueId(uniqueId, folder ? "folder" : "file"),
    cells.get("LastModifiedTime"),
    size,
  );

  return {
    ...base,
    ...(childItemCount !== undefined ? { childItemCount } : {}),
    ...(Object.keys(fields).length > 0 ? { fields } : {}),
  };
}
