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
 * Kiểm tra bit PermissionKind trên SP.BasePermissions — khớp sp.js / PnP `hasPermissions`.
 * Bit index = PermissionKind − 1 (ViewListItems=1 → bit 0 = giá trị 1).
 * @see https://github.com/pnp/pnpjs/blob/main/packages/sp/security/funcs.ts
 */
export function hasPermissionKind(
  permissions: EffectiveBasePermissionsDto,
  kind: PermissionKind,
): boolean {
  if (kind === PermissionKind.EmptyMask) return true;
  if (kind === PermissionKind.FullMask) {
    const low = toWord(permissions.Low);
    const high = toWord(permissions.High);
    return (high & 32767n) === 32767n && low === 65535n;
  }

  const bitIndex = kind - 1;
  const low = toWord(permissions.Low);
  const high = toWord(permissions.High);
  if (bitIndex >= 0 && bitIndex < 32) {
    return (low & (1n << BigInt(bitIndex))) !== 0n;
  }
  if (bitIndex >= 32 && bitIndex < 64) {
    return (high & (1n << BigInt(bitIndex - 32))) !== 0n;
  }
  return false;
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
