# Microsoft Entra delegated permissions (SharePoint REST)

Library dùng **delegated** token của user đang login, audience **SharePoint**.

## Default scopes

```ts
defaultSharePointScopes("https://contoso.sharepoint.com/sites/eOffice")
// → ["https://contoso.sharepoint.com/AllSites.Write"]
```

Entra app registration → API permissions → **SharePoint** → Delegated:

- `AllSites.Read` / `AllSites.Write` (đúng tên trên portal; không phải `AllSites.ReadWrite`)
- Bấm **Grant admin consent**

Đưa scope SharePoint vào MSAL `loginRequest` / silent acquire để tránh consent lần hai.

## App registration checklist

1. SPA redirect URI
2. SharePoint delegated permissions (không chỉ Graph Files.*)
3. Grant admin consent
4. Host truyền `siteUrl` + `tokenProvider` (+ optional `scopes`)

Không cấp client secret vào frontend.
