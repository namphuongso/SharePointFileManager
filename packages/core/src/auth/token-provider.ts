export interface TokenRequest {
  scopes: string[];
  forceRefresh?: boolean;
}

/**
 * Host applications must supply a TokenProvider.
 * The library never logs the user in and never stores tokens.
 */
export interface TokenProvider {
  getAccessToken(request: TokenRequest): Promise<string>;
}

/**
 * Build delegated SharePoint resource scopes for a site URL.
 * Matches Entra SharePoint permission **AllSites.Write** (read + write all site collections).
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aadhttpclient
 */
export function defaultSharePointScopes(siteUrl: string): string[] {
  const origin = new URL(siteUrl).origin;
  return [`${origin}/AllSites.Write`];
}

export const DEFAULT_SHAREPOINT_SCOPE_SUFFIX = "AllSites.Write";
