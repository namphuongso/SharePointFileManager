import { normalizeSiteUrl } from "../utils";
import type { ResolvedSharePointAppConfig, SharePointAppConfig } from "../types/models";

/** Chuẩn hóa siteUrl/siteId cấp app. Thiếu URL thì trả null (UI hiện lỗi, không throw). */
export function resolveAppConfig(
  config: SharePointAppConfig,
): ResolvedSharePointAppConfig | null {
  const siteUrl = config.siteUrl ? normalizeSiteUrl(config.siteUrl) : "";
  if (!siteUrl) return null;
  return {
    ...config,
    siteUrl,
    siteId: config.siteId?.trim() || siteUrl,
  };
}
