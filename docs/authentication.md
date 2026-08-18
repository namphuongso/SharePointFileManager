# Authentication

Mọi project nhúng library đều đã dùng `@azure/msal-browser` / `@azure/msal-react`. Library **reuse session**, không login lần hai.

## Không reuse nguyên access token của app

Token hiện tại thường có audience backend của project (`api://...`). Graph cần token audience `https://graph.microsoft.com`.

Cùng **MSAL instance + account**, gọi `acquireTokenSilent` với Graph scopes.

## TokenProvider

```ts
export interface TokenProvider {
  getAccessToken(request: { scopes: string[]; forceRefresh?: boolean }): Promise<string>;
}
```

Helper:

```ts
const tokenProvider = createMsalTokenProvider({
  instance, // IPublicClientApplication của host
  account,  // accounts[0] hoặc active account
});
```

Bên trong chỉ `acquireTokenSilent`. Không popup, không redirect.

## 401

`GraphClient` gọi lại `getAccessToken({ forceRefresh: true })` một lần rồi retry. Vẫn fail thì UI hiện cần đăng nhập lại — host tự xử lý interactive.

## Consent

Thêm Graph delegated permissions trên **cùng** App Registration. Nên xin scopes lúc login:

```ts
loginRequest: {
  scopes: ["User.Read", "Files.ReadWrite", "Sites.ReadWrite.All", "User.Read.All", "People.Read", "Directory.Read.All"],
}
```

## Không làm

- Client secret / certificate trên FE
- App registration riêng cho library
- Truyền `accessToken: string` tĩnh
- Tạo MSAL instance thứ hai
