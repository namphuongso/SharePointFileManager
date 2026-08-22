import type {
  ResolvedSharePointAppConfig,
  SharePointConfig,
  SharePointLibraryTarget,
} from "../types/models";

/**
 * Gộp cấu hình site/auth cấp app với thư viện đích của từng trang.
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
