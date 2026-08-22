# @namphuongso/sharepoint-file-manager

Thư viện frontend để **đọc và hiển thị** file/folder trên SharePoint Online qua SharePoint REST (`/_api`). Hiện chỉ có browse danh sách — các tính năng khác sẽ dựng lại sau.

```text
Host React app (MSAL)
  → SharePointAppProvider (siteUrl + token)
      → SharePointFileManager libraryName="..."
          → GET /_api → SharePoint Online
```

## Cài đặt

```bash
npm install @namphuongso/sharepoint-file-manager @tanstack/react-query @azure/msal-browser
```

### 1) Provider (app gốc)

```tsx
import "@namphuongso/sharepoint-file-manager/styles.css";
import {
  SharePointAppProvider,
  createMsalTokenProvider,
  defaultSharePointScopes,
} from "@namphuongso/sharepoint-file-manager";

const siteUrl = "https://contoso.sharepoint.com/sites/eOffice";

<SharePointAppProvider
  locale="vi-VN"
  config={{
    siteUrl,
    scopes: defaultSharePointScopes(siteUrl),
    tokenProvider: createMsalTokenProvider({ instance, account }),
  }}
>
  {children}
</SharePointAppProvider>
```

### 2) Trang thư viện

```tsx
<SharePointFileManager libraryName="eDocumentTest" className="h-full" />
```

Library **không login**. Host phải đã đăng nhập Microsoft.

## Entra ID

SharePoint delegated permission `AllSites.Write` + admin consent. Scope: `https://{tenant}.sharepoint.com/AllSites.Write`.

## Docs

- [architecture.md](./docs/architecture.md)
- [authentication.md](./docs/authentication.md)
- [permissions.md](./docs/permissions.md)
- [configuration.md](./docs/configuration.md)
- [rest-api.md](./docs/rest-api.md)
