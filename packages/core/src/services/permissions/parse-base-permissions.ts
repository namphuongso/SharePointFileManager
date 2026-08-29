import { PermissionKind } from "../../types/permissions";
import type { EffectiveBasePermissionsDto, ItemCapabilities } from "../../types/permissions";

/** Chuỗi/số REST → bigint (SP.BasePermissions High/Low). */
function toWord(value: string | number): bigint {
  if (typeof value === "number") return BigInt(value);
  const trimmed = value.trim();
  if (!trimmed) return 0n;
  return BigInt(trimmed);
}

/**
 * Kiểm tra bit PermissionKind trên SP.BasePermissions — tương đương SP.BasePermissions.has().
 * kind < 32 → Low; kind ≥ 32 → High, bit (kind - 32).
 */
export function hasPermissionKind(
  permissions: EffectiveBasePermissionsDto,
  kind: PermissionKind,
): boolean {
  const low = toWord(permissions.Low);
  const high = toWord(permissions.High);
  const word = kind < 32 ? low : high;
  const bit = kind < 32 ? kind : kind - 32;
  return (word & (1n << BigInt(bit))) !== 0n;
}

/** Bitmask Microsoft → capability UX cho file manager. */
export function toItemCapabilities(permissions: EffectiveBasePermissionsDto): ItemCapabilities {
  return {
    canView: hasPermissionKind(permissions, PermissionKind.ViewListItems),
    canAdd: hasPermissionKind(permissions, PermissionKind.AddListItems),
    canEdit: hasPermissionKind(permissions, PermissionKind.EditListItems),
    canDelete: hasPermissionKind(permissions, PermissionKind.DeleteListItems),
    canOpen: hasPermissionKind(permissions, PermissionKind.OpenItems),
    canManagePermissions: hasPermissionKind(permissions, PermissionKind.ManagePermissions),
    raw: permissions,
  };
}
