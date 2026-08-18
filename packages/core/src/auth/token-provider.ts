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

export const DEFAULT_GRAPH_SCOPES = [
  "Files.ReadWrite",
  "Sites.ReadWrite.All",
  "User.Read.All",
  "People.Read",
  "Directory.Read.All",
] as const;
