import type { SharePointItem } from "../types/models";
import { isOfficeFileName } from "./sharepoint-open-url";

export type ItemAction =
  | "open"
  | "openInSharePoint"
  | "download"
  | "preview"
  | "rename"
  | "copy"
  | "move"
  | "share"
  | "manageAccess"
  | "properties"
  | "checkout"
  | "checkin"
  | "discardCheckout"
  | "delete"
  | "versionHistory"
  | "activity";

export function isOfficeOnlineFile(item: SharePointItem): boolean {
  if (item.type !== "file") return false;
  return isOfficeFileName(item.name, item.mimeType);
}

/**
 * Item-level gates from Graph-derived state (checkout, file vs folder).
 * Host FeatureConfig still controls whether the action is offered in the UI.
 */
export function canPerformItemAction(item: SharePointItem, action: ItemAction): boolean {
  const checkedOut = item.capabilities?.isCheckedOut === true;

  switch (action) {
    case "download":
    case "versionHistory":
      return item.type === "file";
    case "preview":
      return item.type === "file" && item.canPreview !== false;
    case "checkout":
      return item.type === "file" && !checkedOut;
    case "checkin":
    case "discardCheckout":
      return item.type === "file" && checkedOut;
    case "open":
      return (
        item.type === "folder" ||
        isOfficeOnlineFile(item) ||
        Boolean(item.webUrl) ||
        item.canPreview !== false
      );
    case "openInSharePoint":
      return Boolean(item.openUrl || item.webUrl);
    case "rename":
    case "delete":
    case "move":
    case "copy":
    case "share":
    case "manageAccess":
    case "properties":
    case "activity":
      return true;
    default:
      return true;
  }
}
