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
  if (BUILTIN_COLUMN_NAMES.has(key)) return true;
  if (key.startsWith("_")) return true;
  if (key.startsWith("OData__")) return true;
  if (key.endsWith("LookupId")) return true;
  if (key.endsWith("StringId")) return true;
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
  if (typeof value === "object") {
    const record = value as Record<string, unknown>;
    if (typeof record.displayName === "string") return record.displayName;
    if (typeof record.Label === "string") return record.Label;
    if (typeof record.lookupValue === "string") return record.lookupValue;
    if (Array.isArray(value)) {
      const labels = value
        .map((entry) => {
          if (typeof entry === "string") return entry;
          if (entry && typeof entry === "object" && typeof (entry as Record<string, unknown>).displayName === "string") {
            return (entry as Record<string, string>).displayName;
          }
          return undefined;
        })
        .filter((entry): entry is string => Boolean(entry));
      if (labels.length > 0) return labels.join(", ");
    }
  }
  return String(value);
}
