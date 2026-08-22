import type { TokenProvider } from "@namphuongso/sharepoint-file-manager-core";
import type { CreateMsalTokenProviderOptions } from "../types";

/**
 * Adapter auth: host đã đăng nhập MSAL, truyền PublicClientApplication + account đang dùng.
 * Thư viện chỉ gọi acquireTokenSilent.
 */
export function createMsalTokenProvider(
  options: CreateMsalTokenProviderOptions,
): TokenProvider {
  return {
    async getAccessToken({ scopes, forceRefresh }) {
      const result = await options.instance.acquireTokenSilent({
        account: options.account,
        scopes,
        forceRefresh,
      } as never);
      return result.accessToken;
    },
  };
}
