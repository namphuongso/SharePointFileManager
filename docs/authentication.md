# Authentication

Library **không login**. Host đã đăng nhập MSAL; library chỉ xin token SharePoint silent.

## Audience

Token backend nội bộ (`api://...`) **không** gọi được SharePoint REST.

Cùng **MSAL instance + account**, gọi `acquireTokenSilent` với SharePoint scopes:

```ts
scopes: defaultSharePointScopes("https://contoso.sharepoint.com/sites/eOffice")
// → ["https://contoso.sharepoint.com/AllSites.Write"]
```

## TokenProvider

```ts
createMsalTokenProvider({ instance, account })
```

Khi REST trả 401, client gọi lại `getAccessToken({ forceRefresh: true })` một lần rồi retry. Vẫn fail thì UI hiện cần đăng nhập lại — host tự xử lý interactive.

## Entra

Trên cùng App Registration (SPA), thêm SharePoint delegated **`AllSites.Write`**, admin consent, và đưa scope vào `loginRequest`.

Xem [permissions.md](./permissions.md).
