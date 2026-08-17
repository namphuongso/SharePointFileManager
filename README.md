# @namphuongso/sharepoint-file-manager

Thư viện frontend dùng chung để quản lý file/folder trên **SharePoint Online** qua **Microsoft Graph**. Không có backend trung gian, không có database quyền riêng. SharePoint là source of truth.

```text
Host React app (MSAL)
  → @namphuongso/sharepoint-file-manager
      → SharePoint core (Graph client + services)
          → Microsoft Graph
              → SharePoint Online
```

## Cài đặt

```bash
npm install @namphuongso/sharepoint-file-manager @tanstack/react-query @azure/msal-browser
```

```ts
import "@namphuongso/sharepoint-file-manager/styles.css";
import { SharePointFileManager, createMsalTokenProvider } from "@namphuongso/sharepoint-file-manager";
import { useMsal } from "@azure/msal-react";

export function Documents() {
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  if (!account) return null;

  const tokenProvider = createMsalTokenProvider({ instance, account });

  return (
    <SharePointFileManager
      locale="vi-VN"
      config={{
        siteId: "contoso.sharepoint.com,site-guid,web-guid",
        driveId: "b!...",
        tokenProvider,
      }}
    />
  );
}
```

Library **không login**. App host phải đã đăng nhập Microsoft. `createMsalTokenProvider` chỉ gọi `acquireTokenSilent` để lấy token Graph.

## Entra ID

Trên App Registration **hiện có** của project (SPA), thêm delegated permissions:

- `Files.ReadWrite`
- `Sites.ReadWrite.All` (cần cho share/manage access trên document library)

Admin consent. Nên đưa Graph scopes vào `loginRequest` của MSAL để tránh consent lần hai.

Không đưa client secret / certificate vào frontend.

## Feature flags

```ts
features: {
  delete: false,
  share: true,
}
```

Nút vẫn bị ẩn/disable theo flag, nhưng quyền thật do SharePoint quyết định. Graph trả 403 thì UI hiện trạng thái không có quyền.

## Tài liệu

- [Architecture](docs/architecture.md)
- [Authentication](docs/authentication.md)
- [Configuration](docs/configuration.md)
- [Graph API map](docs/graph-api.md)
- [Entra permissions](docs/permissions.md)
- [Troubleshooting](docs/troubleshooting.md)

## Monorepo

```bash
npm install
npm test
npm run build
npm run dev:example
```
