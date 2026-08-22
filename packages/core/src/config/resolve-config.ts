import { defaultSharePointScopes } from "../auth/token-provider";
import type { ResolvedSharePointConfig, SharePointConfig } from "../types/models";
import { resolveAppConfig } from "./resolve-app-config";

/** Chuẩn hóa siteUrl, siteId, rootItemId, scopes trước khi tạo client. */

export function resolveConfig(config: SharePointConfig): ResolvedSharePointConfig {
  const app = resolveAppConfig(config);
  if (!app) {
    throw new Error("SharePointConfig.siteUrl is required (SharePoint REST base URL)");
  }
  if (!config.tokenProvider) {
    throw new Error("SharePointConfig.tokenProvider is required");
  }

  return {
    ...config,
    ...app,
    rootItemId: config.rootItemId ?? "root",
    scopes: config.scopes?.length ? config.scopes : defaultSharePointScopes(app.siteUrl),
  };
}
