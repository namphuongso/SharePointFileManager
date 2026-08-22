# Cấu trúc thư viện (cho dev)

Một file này mô tả cách tổ chức code để phát triển và mở rộng. Cách dùng (cài đặt, Entra, JSX) nằm ở [README](../README.md).

## Vai trò

Frontend **chỉ đọc** file/folder SharePoint Online qua REST `GET /_api`. Host app đã đăng nhập MSAL; thư viện không login.

```text
Host (MSAL)
  → packages/react  SharePointAppProvider + SharePointFileManager
      → packages/core  SharePointClient
          → SharePointRestClient.get
              → SharePoint Online
```

UI không `fetch` SharePoint trực tiếp.

## Packages

| Path | Việc |
|---|---|
| `packages/core` | Token, REST GET, resolve library, list folder/file |
| `packages/react` | Provider, hook, bảng danh sách, i18n, Fluent |
| `examples/react-vite` | App mẫu |

Hiện **không** có upload, share, checkout, search, Graph.

## Core — thứ tự đọc

```text
types/                      models, rest, errors
client.ts                   Ghép config + REST + cache thư viện + FolderService
auth/                       defaultSharePointScopes
config/                     resolve-config, create-sharepoint-config
rest/                       client GET, build-url, parse-body, throttle, odata
services/library/           resolve theo libraryName
services/folder/            listChildren + resolve path
mappers/                    SP.File / SP.Folder → SharePointItem
errors/                     SharePointError + map HTTP/OData
```

Luồng list một folder:

1. `getLibrary()` (cache) → `rootFolderServerRelativeUrl`
2. Nếu không phải root: `GetFolderById` lấy `ServerRelativeUrl`
3. `.../Folders` và `.../Files`
4. Bỏ folder `Forms`, map, sort folder trước file

Thêm cột UI: sửa `$select` trong `services/folder/folder.ts` rồi map trong `mappers/rest-item.ts`.

## React — thứ tự đọc

```text
types/                      Props, Messages, MSAL
auth/                       createMsalTokenProvider
provider/                   AppProvider + SharePointProvider + hooks
hooks/                      useFolderChildren, getErrorMessage
fluent/                     theme, isDarkTheme
i18n/                       messages
utils/                      format bytes / ngày
components/file-manager/    UI duyệt thư viện (browser, bảng, empty, banner, icons)
```

## Quy tắc khi sửa

- Interface/type nằm ở `types/` — không khai báo chung file logic.
- Việc liên quan nằm **cùng thư mục**; mỗi file một việc.
- REST chỉ đọc: thêm ghi/xóa thì tách service mới, không nhồi vào `services/library`.
- `$select` chỉ field đang dùng; quan hệ (`Author`, `ListItemAllFields`) cần `$expand`.
- Comment theo **khối / hàm** (ý định + vì sao), tiếng Việt, không comment từng dòng gán biến.
