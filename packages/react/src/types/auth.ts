export type SilentTokenResult = { accessToken: string };

/** Kiểu tối thiểu của instance MSAL (tương thích lệch type IPublicClientApplication). */
export type MsalSilentTokenSource = {
  acquireTokenSilent: (request: never) => Promise<SilentTokenResult>;
};

export interface CreateMsalTokenProviderOptions {
  instance: MsalSilentTokenSource;
  account: object;
}
