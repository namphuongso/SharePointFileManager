import type {
  ResolvedSharePointAppConfig,
  SharePointConfig,
  SharePointLibraryTarget,
} from "../types/models";

/**
 * Merge app-level site/auth settings with a per-route library target.
 *
 * ```ts
 * const app = { siteId, tokenProvider, features };
 * const config = createSharePointConfig(app, { libraryName: "eDocumentTest" });
 * ```
 */
export function createSharePointConfig(
  app: ResolvedSharePointAppConfig,
  target: SharePointLibraryTarget = {},
): SharePointConfig {
  if (!app.siteId) {
    throw new Error("createSharePointConfig requires app.siteId (resolve siteUrl via SharePointAppProvider first)");
  }
  return {
    ...app,
    siteId: app.siteId,
    ...target,
  };
}
