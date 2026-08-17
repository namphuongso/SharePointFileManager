import type { TokenProvider } from "@namphuongso/sharepoint-file-manager-core";

type SilentTokenResult = { accessToken: string };

/**
 * Structural MSAL instance. `never` on the request param keeps this compatible
 * with msal-browser 3.x IPublicClientApplication type drift.
 */
export type MsalSilentTokenSource = {
  acquireTokenSilent: (request: never) => Promise<SilentTokenResult>;
};

/**
 * Production auth adapter. Host apps already signed in with MSAL pass their
 * PublicClientApplication + active account. The library only calls acquireTokenSilent.
 */
export function createMsalTokenProvider(options: {
  instance: MsalSilentTokenSource;
  account: object;
}): TokenProvider {
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
