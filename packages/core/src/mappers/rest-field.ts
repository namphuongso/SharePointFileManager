import type { SharePointField } from "../types/models";
import type { RestField } from "../types/rest";

/**
 * JSON SP.Field → SharePointField (option ẩn/hiện cột).
 * Thiếu InternalName thì bỏ — không có key ổn định.
 */
export function mapRestField(field: RestField): SharePointField | undefined {
  const internalName = field.InternalName?.trim();
  if (!internalName) return undefined;

  return {
    id: field.Id,
    title: field.Title?.trim() || internalName,
    internalName,
    entityPropertyName: field.EntityPropertyName,
    typeAsString: field.TypeAsString,
    fieldTypeKind: field.FieldTypeKind,
    hidden: field.Hidden === true,
    readOnly: field.ReadOnlyField === true,
    required: field.Required === true,
    sortable: field.Sortable,
    filterable: field.Filterable,
  };
}
