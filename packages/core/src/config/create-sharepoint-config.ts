import type {
  ResolvedSharePointAppConfig,
  SharePointConfig,
  SharePointLibraryTarget,
} from "../types/models";

/**
 * Merge app-level site/auth settings with a per-route library target.
 */
export function createSharePointConfig(
  app: ResolvedSharePointAppConfig,
  target: SharePointLibraryTarget = {},
): SharePointConfig {
  if (!app.siteUrl) {
    throw new Error("createSharePointConfig requires app.siteUrl");
  }
  return {
    ...app,
    siteUrl: app.siteUrl,
    siteId: app.siteId || app.siteUrl,
    ...target,
  };
}
