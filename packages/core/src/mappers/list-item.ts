import type { ListColumn, ListItemFields } from "../types/models";

interface GraphListColumn {
  id?: string;
  name?: string;
  displayName?: string;
  hidden?: boolean;
  readOnly?: boolean;
  text?: unknown;
  choice?: unknown;
  dateTime?: unknown;
  number?: unknown;
  boolean?: unknown;
  lookup?: unknown;
  personOrGroup?: unknown;
}

interface GraphListItem {
  id?: string;
  contentType?: { id?: string; name?: string };
  fields?: Record<string, unknown>;
}

const BUILTIN_COLUMN_NAMES = new Set([
  "ContentType",
  "ContentTypeId",
  "FileLeafRef",
  "FileRef",
  "FileDirRef",
  "File_x0020_Type",
  "File_x0020_Size",
  "FSObjType",
  "Modified",
  "Created",
  "Author",
  "Editor",
  "AuthorLookupId",
  "EditorLookupId",
  "Created_x0020_Date",
  "Modified_x0020_By",
  "_EditMenuTableStart",
  "_EditMenuTableEnd",
  "_CopySource",
  "GUID",
  "ID",
  "ItemChildCount",
  "FolderChildCount",
  "UniqueId",
  "DocIcon",
  "_UIVersionString",
  "_UIVersion",
  "_Level",
  "_IsCurrentVersion",
  "_HasCopyDestinations",
  "_ModerationStatus",
  "_ModerationComments",
  "_ComplianceFlags",
  "_ComplianceTag",
  "_ComplianceTagWrittenTime",
  "_ComplianceTagUserId",
  "_DisplayName",
  "SelectTitle",
  "SelectFilename",
  "SelectPath",
  "AppAuthor",
  "AppEditor",
  "Edit",
  "ParentVersionStringLookupId",
  "CheckoutUserId",
  "CheckedOutUserId",
  "CheckoutUser",
  "LinkFilename",
  "LinkFilenameNoMenu",
]);

export function mapListColumn(column: GraphListColumn): ListColumn | undefined {
  if (!column.id || !column.name) return undefined;
  const displayName = typeof column.displayName === "string" ? column.displayName.trim() : column.name;
  return {
    id: column.id,
    name: column.name,
    displayName,
    readOnly: column.readOnly === true,
    hidden: column.hidden === true,
    type: detectColumnType(column),
  };
}

export function isVisibleListColumn(column: ListColumn): boolean {
  if (column.hidden) return false;
  if (isSystemMetadataKey(column.name)) return false;
  return true;
}

export function mapListItemFields(itemId: string, listItem?: GraphListItem): ListItemFields {
  const raw = listItem?.fields ?? {};
  const fields: Record<string, string | number | boolean | null> = {};
  for (const [key, value] of Object.entries(raw)) {
    const normalized = normalizeFieldValue(value);
    if (normalized !== undefined) fields[key] = normalized;
  }
  return {
    itemId,
    listItemId: listItem?.id,
    contentType: listItem?.contentType?.name,
    fields,
  };
}

export function extractMetadataFromListItem(listItem?: GraphListItem): {
  listItemId?: string;
  contentType?: string;
  metadata?: Record<string, string | number | boolean | null>;
} {
  if (!listItem) return {};
  const mapped = mapListItemFields("", listItem);
  const metadata: Record<string, string | number | boolean | null> = {};
  for (const [key, value] of Object.entries(mapped.fields)) {
    if (isSystemMetadataKey(key)) continue;
    metadata[key] = value;
  }
  return {
    listItemId: listItem.id,
    contentType: listItem.contentType?.name,
    metadata: Object.keys(metadata).length > 0 ? metadata : undefined,
  };
}

function isSystemMetadataKey(key: string): boolean {
  const normalized = key.trim();
  if (normalized.startsWith("@")) return true;
  if (normalized.toLowerCase() === "id") return true;
  if (BUILTIN_COLUMN_NAMES.has(normalized)) return true;
  // Keep custom internal names that begin with "_" and avoid dropping valid metadata.
  if (key.startsWith("OData__")) return true;
  return false;
}

function detectColumnType(column: GraphListColumn): string | undefined {
  if (column.text) return "text";
  if (column.choice) return "choice";
  if (column.dateTime) return "dateTime";
  if (column.number) return "number";
  if (column.boolean) return "boolean";
  if (column.lookup) return "lookup";
  if (column.personOrGroup) return "personOrGroup";
  return undefined;
}

function normalizeFieldValue(value: unknown): string | number | boolean | null | undefined {
  if (value === null || value === undefined) return null;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  if (Array.isArray(value)) return normalizeArrayFieldValue(value);
  if (typeof value === "object") return normalizeObjectFieldValue(value);
  return String(value);
}

function normalizeArrayFieldValue(values: unknown[]): string | null {
  const labels = values
    .map((entry) => normalizeFieldValue(entry))
    .filter((entry): entry is string | number | boolean => entry !== null && entry !== undefined)
    .map((entry) => String(entry).trim())
    .filter((entry) => entry.length > 0);
  if (labels.length === 0) return null;
  return [...new Set(labels)].join(", ");
}

function normalizeObjectFieldValue(value: object): string {
  const record = value as Record<string, unknown>;
  const candidates = [
    record.LookupValue,
    record.lookupValue,
    record.displayName,
    record.DisplayName,
    record.Title,
    record.Name,
    record.name,
    record.Label,
    record.label,
    record.Email,
    record.email,
    record.Url,
    record.value,
  ];
  for (const candidate of candidates) {
    if (typeof candidate === "string" && candidate.trim()) return candidate.trim();
    if (typeof candidate === "number" || typeof candidate === "boolean") return String(candidate);
  }

  if (Array.isArray(record.results)) {
    const normalized = normalizeArrayFieldValue(record.results);
    if (normalized) return normalized;
  }

  return JSON.stringify(record);
}
