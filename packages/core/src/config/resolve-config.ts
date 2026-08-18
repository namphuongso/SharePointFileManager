import { DEFAULT_GRAPH_SCOPES } from "../auth/token-provider";
import type {
  FeatureConfig,
  ResolvedSharePointConfig,
  SharePointConfig,
} from "../types/models";

export const DEFAULT_FEATURES: Required<FeatureConfig> = {
  upload: true,
  download: true,
  createFolder: true,
  rename: true,
  delete: true,
  copy: true,
  move: true,
  share: true,
  manageAccess: true,
  search: true,
  preview: true,
  versionHistory: true,
  openInSharePoint: true,
  properties: true,
  checkout: true,
  createOfficeFile: true,
  globalSearch: true,
  metadata: true,
  activityLog: true,
  infiniteScroll: true,
  dragDropMove: true,
  bulkMetadata: true,
  copyProgress: true,
  enableDeltaSync: false,
  enableAnalytics: false,
  enableActivities: true,
};

export function resolveConfig(config: SharePointConfig): ResolvedSharePointConfig {
  if (!config.siteId) {
    throw new Error("SharePointConfig.siteId is required");
  }
  if (!config.tokenProvider) {
    throw new Error("SharePointConfig.tokenProvider is required");
  }

  return {
    ...config,
    rootItemId: config.rootItemId ?? "root",
    scopes: config.scopes ?? [...DEFAULT_GRAPH_SCOPES],
    graphBaseUrl: config.graphBaseUrl ?? "https://graph.microsoft.com/v1.0",
    features: { ...DEFAULT_FEATURES, ...config.features },
  };
}

export function isFeatureEnabled(
  config: ResolvedSharePointConfig,
  feature: keyof FeatureConfig,
): boolean {
  return config.features[feature] !== false;
}
