import { defaultSharePointScopes } from "../auth/token-provider";
import type { ResolvedSharePointConfig, SharePointConfig } from "../types/models";

export function resolveConfig(config: SharePointConfig): ResolvedSharePointConfig {
  const siteUrl = config.siteUrl?.trim();
  if (!siteUrl) {
    throw new Error("SharePointConfig.siteUrl is required (SharePoint REST base URL)");
  }
  if (!config.tokenProvider) {
    throw new Error("SharePointConfig.tokenProvider is required");
  }

  const normalizedUrl = siteUrl.replace(/\/$/, "");
  return {
    ...config,
    siteUrl: normalizedUrl,
    siteId: config.siteId?.trim() || normalizedUrl,
    rootItemId: config.rootItemId ?? "root",
    scopes: config.scopes?.length ? config.scopes : defaultSharePointScopes(normalizedUrl),
  };
}
