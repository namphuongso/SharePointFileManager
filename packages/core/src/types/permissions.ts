import type { SharePointItemType } from "./models";

/**
 * Giá trị PermissionKind — khớp Microsoft.SharePoint.Client.PermissionKind (dùng làm chỉ số bit).
 * @see https://learn.microsoft.com/en-us/dotnet/api/microsoft.sharepoint.client.permissionkind
 */
export enum PermissionKind {
  EmptyMask = 0,
  ViewListItems = 1,
  AddListItems = 2,
  EditListItems = 3,
  DeleteListItems = 4,
  ApproveItems = 5,
  OpenItems = 6,
  ViewVersions = 7,
  DeleteVersions = 8,
  CancelCheckout = 9,
  ManagePersonalViews = 10,
  ManageLists = 12,
  ViewFormPages = 13,
  AnonymousSearchAccessList = 14,
  Open = 17,
  ViewPages = 18,
  AddAndCustomizePages = 19,
  ApplyThemeAndBorder = 20,
  ApplyStyleSheets = 21,
  ViewUsageData = 22,
  CreateSSCSite = 23,
  ManageSubwebs = 24,
  CreateGroups = 25,
  ManagePermissions = 26,
  BrowseDirectories = 27,
  BrowseUserInfo = 28,
  AddDelPrivateWebParts = 29,
  UpdatePersonalWebParts = 30,
  ManageWeb = 31,
  AnonymousSearchAccessWebLists = 32,
  UseClientIntegration = 37,
  UseRemoteAPIs = 38,
  ManageAlerts = 39,
  CreateAlerts = 40,
  EditMyUserInfo = 41,
  EnumeratePermissions = 63,
  FullMask = 65,
}

/** REST EffectiveBasePermissions — SP.BasePermissions (High/Low bitmask). */
export interface EffectiveBasePermissionsDto {
  High: string | number;
  Low: string | number;
}

/**
 * Quyền UX map từ PermissionKind (Microsoft) — không phải REST response riêng.
 * Action mới: thêm boolean tương ứng PermissionKind trong toItemCapabilities.
 */
export interface ItemCapabilities {
  canView: boolean;
  canAdd: boolean;
  canEdit: boolean;
  canDelete: boolean;
  canOpen: boolean;
  canManagePermissions: boolean;
  raw: EffectiveBasePermissionsDto;
}

export type PermissionItemType = SharePointItemType;
