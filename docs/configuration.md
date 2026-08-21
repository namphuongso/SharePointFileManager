# Configuration

## Two places only (host app)

| Place | What |
|---|---|
| **1. `SharePointAppProvider`** | `siteUrl` or `siteId`, `tokenProvider`, optional `features` / scopes |
| **2. `SharePointFileManager`** | `libraryName` (or `listId` / `driveId`) |

Everything else (resolve site, Graph client, UI, hooks) lives inside the library.

```ts
import {
  SharePointAppProvider,
  SharePointFileManager,
  createMsalTokenProvider,
} from "@namphuongso/sharepoint-file-manager";

// 1) App root
<SharePointAppProvider
  locale="vi-VN"
  config={{
    siteUrl: "https://contoso.sharepoint.com/sites/eOffice",
    tokenProvider,
    features: { delete: false },
  }}
>
  <AppRoutes />
</SharePointAppProvider>

// 2) Feature route — chỉ tên thư viện
<SharePointFileManager libraryName="eDocumentTest" />
```

`siteUrl` được provider resolve sang Graph `siteId` tự động. Có thể truyền sẵn `siteId` nếu đã biết.

`embedded` mặc định `true` khi dùng AppProvider + `libraryName` (ẩn header + navigation của `SharePointShell`). Standalone full shell: truyền `config` đầy đủ và `embedded={false}`.

## SharePointAppConfig

```ts
interface SharePointAppConfig {
  siteId?: string;   // Graph site id
  siteUrl?: string;  // web URL — provider resolves → siteId
  scopes?: string[];
  graphBaseUrl?: string;
  tokenProvider: TokenProvider;
  features?: FeatureConfig;
}

interface SharePointLibraryTarget {
  libraryName?: string;  // resolve drive theo tên (vd. "eDocumentTest")
  listId?: string;
  driveId?: string;
  rootItemId?: string;   // mặc định "root"
}
```

`tenantId` / `clientId` không bắt buộc khi nhúng app thật — MSAL host đã có.

## Multi-module

```ts
<SharePointFileManager libraryName="eDocumentTest" />
<SharePointFileManager libraryName="ISO Documents" rootItemId="01ABC..." />
```

## Features

Mọi flag đều ánh xạ chức năng dựa trên endpoint Microsoft Graph `v1.0`.

Action trên UI = feature flag AND kết quả Graph. 403 → permission denied.

| Flag | Mặc định | Việc |
|---|---|---|
| `upload` | true | Tải file / thư mục lên |
| `download` | true | Tải xuống |
| `createFolder` | true | Tạo thư mục |
| `rename` | true | Đổi tên |
| `delete` | true | Xóa (recycle bin) |
| `copy` | true | Sao chép |
| `move` | true | Di chuyển |
| `share` | true | Invite / create link |
| `manageAccess` | true | Xem / sửa / gỡ permission |
| `search` | true | Search trong folder |
| `globalSearch` | true | Search toàn library (`/search/query`) + filters |
| `preview` | true | Preview file |
| `versionHistory` | true | Phiên bản / restore |
| `openInSharePoint` | true | Mở trong SharePoint / Office |
| `properties` | true | Details pane |
| `checkout` | true | Checkout / checkin / discard |
| `metadata` | true | Cột list / fields |
| `bulkMetadata` | true | Sửa metadata hàng loạt |
| `activityLog` | true | Item activities |
| `infiniteScroll` | true | Phân trang children |
| `dragDropMove` | true | Kéo thả di chuyển |
| `copyProgress` | true | Progress khi copy |

Default scopes hiện tại:

```text
Files.ReadWrite
Sites.ReadWrite.All
User.Read.All
People.Read
Directory.Read.All
```

Khi `.../root/children` rỗng vì user chỉ có quyền item-level, UI fallback sang `GET /sites/{site}/lists/{list}/items?$expand=driveItem` (cùng nguồn list view SharePoint).
