# @namphuongso/sharepoint-file-manager

Thư viện frontend dùng chung để quản lý file/folder trên **SharePoint Online** qua **Microsoft Graph**. Không có backend trung gian, không có database quyền riêng. SharePoint là source of truth.

```text
Host React app (MSAL)
  → SharePointAppProvider (config: siteUrl/siteId + token + features)
      → SharePointFileManager libraryName="..."
          → Microsoft Graph → SharePoint Online
```

Host chỉ cần **2 chỗ**. Mọi resolve site, Graph client, UI, hooks nằm trong thư viện.

## Cài đặt

```bash
npm install @namphuongso/sharepoint-file-manager @tanstack/react-query @azure/msal-browser
```

### 1) Provider + config (app gốc)

```tsx
import "@namphuongso/sharepoint-file-manager/styles.css";
import {
  SharePointAppProvider,
  SharePointFileManager,
  createMsalTokenProvider,
} from "@namphuongso/sharepoint-file-manager";
import { useMsal } from "@azure/msal-react";

function AppRoot({ children }: { children: React.ReactNode }) {
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  if (!account) return <>{children}</>;

  return (
    <SharePointAppProvider
      locale="vi-VN"
      config={{
        siteUrl: "https://contoso.sharepoint.com/sites/eOffice",
        // hoặc siteId: "contoso.sharepoint.com,site-guid,web-guid",
        tokenProvider: createMsalTokenProvider({ instance, account }),
        features: { delete: false }, // optional
      }}
    >
      {children}
    </SharePointAppProvider>
  );
}
```

### 2) Nơi sử dụng — chỉ `libraryName`

```tsx
export function DocumentTestPage() {
  return <SharePointFileManager libraryName="eDocumentTest" className="h-full" />;
}
```

### Standalone (không AppProvider)

```tsx
<SharePointFileManager
  locale="vi-VN"
  embedded={false}
  config={{ siteId, driveId, tokenProvider }}
/>
```

Library **không login**. App host phải đã đăng nhập Microsoft. `createMsalTokenProvider` chỉ gọi `acquireTokenSilent` để lấy token Graph.

## Entra ID

Trên App Registration **hiện có** của project (SPA), thêm delegated permissions:

- `Files.ReadWrite`
- `Sites.ReadWrite.All` (cần cho share/manage access trên document library)
- `People.Read`, `User.Read.All`, `Directory.Read.All` (people picker khi share)

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
