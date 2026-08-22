import { normalizeSiteUrl } from "./site-url";

/** nextLink tuyệt đối, hoặc ghép vào site nếu SharePoint trả path. */
export function resolveODataNextLink(nextLink: string | undefined, siteUrl: string): string | undefined {
  if (!nextLink) return undefined;
  if (/^https?:\/\//i.test(nextLink)) return nextLink;
  const site = normalizeSiteUrl(siteUrl);
  if (nextLink.startsWith("/")) return `${new URL(site).origin}${nextLink}`;
  return `${site}/${nextLink}`;
}
