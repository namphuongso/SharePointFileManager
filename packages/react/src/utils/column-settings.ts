export type BuiltInColumnId = "modified" | "modifiedBy" | "size" | "created" | "createdBy";

export interface ColumnVisibilitySettings {
  modified: boolean;
  modifiedBy: boolean;
  size: boolean;
  created: boolean;
  createdBy: boolean;
  /** Internal column names for visible custom metadata columns. */
  metadataColumnNames: string[];
}

export const DEFAULT_COLUMN_VISIBILITY: ColumnVisibilitySettings = {
  modified: true,
  modifiedBy: true,
  size: false,
  created: false,
  createdBy: false,
  metadataColumnNames: [],
};

const STORAGE_PREFIX = "spm-columns:";

export function loadColumnSettings(scope: string): ColumnVisibilitySettings {
  if (typeof localStorage === "undefined") return DEFAULT_COLUMN_VISIBILITY;
  try {
    const raw = localStorage.getItem(`${STORAGE_PREFIX}${scope}`);
    if (!raw) return DEFAULT_COLUMN_VISIBILITY;
    const parsed = JSON.parse(raw) as Partial<ColumnVisibilitySettings>;
    return {
      ...DEFAULT_COLUMN_VISIBILITY,
      ...parsed,
      metadataColumnNames: Array.isArray(parsed.metadataColumnNames) ? parsed.metadataColumnNames : [],
    };
  } catch {
    return DEFAULT_COLUMN_VISIBILITY;
  }
}

export function saveColumnSettings(scope: string, settings: ColumnVisibilitySettings): void {
  if (typeof localStorage === "undefined") return;
  localStorage.setItem(`${STORAGE_PREFIX}${scope}`, JSON.stringify(settings));
}

export function defaultMetadataColumnNames(columnNames: string[], limit = 4): string[] {
  return columnNames.slice(0, limit);
}
