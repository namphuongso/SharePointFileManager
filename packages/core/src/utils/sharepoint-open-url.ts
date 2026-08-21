import type { SharePointItem } from "../types/models";

export interface GraphSharePointIds {
  listId?: string;
  listItemId?: string;
  listItemUniqueId?: string;
  siteId?: string;
  siteUrl?: string;
  tenantId?: string;
  webId?: string;
}

export interface GraphOpenUrlItem {
  name?: string;
  webUrl?: string;
  file?: { mimeType?: string };
  folder?: unknown;
  sharepointIds?: GraphSharePointIds;
}

/** SharePoint web UI opens Office files via Doc.aspx, not the raw library path. */
export function buildSharePointDocOpenUrl(
  siteUrl: string,
  listItemUniqueId: string,
  fileName: string,
): string {
  const base = siteUrl.replace(/\/$/, "");
  const guid = listItemUniqueId.replace(/^\{|\}$/g, "").trim();
  const sourcedoc = encodeURIComponent(`{${guid}}`);
  const file = encodeURIComponent(fileName);
  return `${base}/_layouts/15/Doc.aspx?sourcedoc=${sourcedoc}&file=${file}&action=default&mobileredirect=true`;
}

export function isOfficeFileName(name: string, mimeType?: string): boolean {
  const lowerName = name.toLowerCase();
  const mime = mimeType?.toLowerCase() ?? "";
  return (
    /\.(doc|docx|dotx|odt|xls|xlsx|xlsm|csv|ods|ppt|pptx|ppsx|odp)$/.test(lowerName) ||
    mime.includes("word") ||
    mime.includes("excel") ||
    mime.includes("spreadsheet") ||
    mime.includes("powerpoint") ||
    mime.includes("presentation") ||
    mime.includes("ms-word") ||
    mime.includes("officedocument")
  );
}

export function resolveGraphItemOpenUrl(item: GraphOpenUrlItem): string | undefined {
  if (item.folder) return item.webUrl;
  if (!item.name || !item.file) return item.webUrl;

  const siteUrl = item.sharepointIds?.siteUrl;
  const uniqueId = item.sharepointIds?.listItemUniqueId;
  if (siteUrl && uniqueId && isOfficeFileName(item.name, item.file.mimeType)) {
    return buildSharePointDocOpenUrl(siteUrl, uniqueId, item.name);
  }
  return item.webUrl;
}

/** Prefer Doc.aspx open URL; fall back to Graph webUrl. */
export function resolveItemOpenUrl(item: SharePointItem): string | undefined {
  if (item.type === "folder") return item.webUrl;
  return item.openUrl ?? item.webUrl;
}

export function isDirectFileDownloadUrl(url: string): boolean {
  try {
    const parsed = new URL(url);
    if (parsed.pathname.includes("/download.aspx")) return true;
    if (parsed.searchParams.get("download") === "1") return true;
    return /\.(docx?|xlsx?|xlsm|pptx?|ppsx|csv|odt|ods|odp)$/i.test(parsed.pathname);
  } catch {
    return /\.(docx?|xlsx?|xlsm|pptx?|ppsx|csv|odt|ods|odp)(\?|$)/i.test(url);
  }
}
