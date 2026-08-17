export function encodePathSegment(name: string): string {
  return encodeURIComponent(name);
}

/**
 * Graph site resource path.
 *
 * Hostname/path identifiers must keep a trailing colon before nested resources:
 * `/sites/{hostname}:/{server-relative-path}:/drives`
 *
 * GUID and composite IDs use a normal slash:
 * `/sites/{siteId}/drives`
 */
export function siteResourcePath(siteId: string, resource = ""): string {
  const id = siteId.trim().replace(/^\/+|\/+$/g, "");
  const suffix = resource.replace(/^\/+/, "");
  if (!id) {
    throw new Error("SharePoint siteId is required");
  }

  if (id.includes(":/")) {
    const withColon = id.endsWith(":") ? id : `${id}:`;
    return suffix ? `/sites/${withColon}/${suffix}` : `/sites/${id}`;
  }

  return suffix ? `/sites/${id}/${suffix}` : `/sites/${id}`;
}

export function itemUrl(driveId: string, itemId: string): string {
  return `/drives/${driveId}/items/${itemId}`;
}

export function childrenUrl(driveId: string, itemId: string): string {
  return `${itemUrl(driveId, itemId)}/children`;
}

export function pathByNameUrl(
  driveId: string,
  parentId: string,
  fileName: string,
  suffix: string,
): string {
  return `/drives/${driveId}/items/${parentId}:/${encodePathSegment(fileName)}:/${suffix}`;
}
