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

Toolbar có sẵn menu **VI/EN**. Khi người dùng đổi ngôn ngữ, thư viện sẽ gọi lại fields
với header `Accept-Language` mới. Nếu host muốn tự kiểm soát, `locale` vẫn là prop động.

```tsx
const [locale, setLocale] = useState("vi-VN");

<SharePointFileManager
  libraryName="eDocumentTest"
  locale={locale}
  className="h-full"
/>
```

Muốn ẩn menu tích hợp sẵn:

```tsx
<SharePointFileManager
  libraryName="eDocumentTest"
  showLanguageSwitcher={false}
/>
```

Library **không login**. Host phải đã đăng nhập Microsoft.

## Entra ID

SharePoint delegated permission `AllSites.Write` + admin consent. Scope: `https://{tenant}.sharepoint.com/AllSites.Write`.

## Cấu trúc code

Xem [docs/LIBRARY.md](./docs/LIBRARY.md) — một file cho dev (package, luồng REST, chỗ sửa `$select`).
