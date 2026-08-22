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
client.ts                   Ghép config + REST + cache thư viện + FolderService + FieldService
auth/                       defaultSharePointScopes
config/                     resolve-config, create-sharepoint-config
rest/                       client GET, build-url, parse-body, throttle, odata
services/library/           resolve theo libraryName
services/folder/            listChildren + resolve path
services/fields/            list schema cột (GET .../fields, không $select)
mappers/                    list item / SP.Field → model
errors/                     SharePointError + map HTTP/OData
```

Luồng list một folder:

1. `getLibrary()` (cache) → `listId` + `rootFolderServerRelativeUrl`
2. Nếu không phải root: `GetFolderById` lấy `ServerRelativeUrl` (= FileDirRef)
3. `GET web/lists(guid)/items?$filter=FileDirRef eq '...'` — `$select=*` + property `$expand`; `$expand=File,Folder,Author,Editor`, `$top=30`
4. Trang sau: GET nguyên `@odata.nextLink` (`SharePointRestClient.getUrl`)
5. Bỏ folder `Forms`, map `File.UniqueId` / `Folder.UniqueId`

Schema cột (option ẩn/hiện): `GET web/lists(guid'{listId}')/fields` — không `$select`.

Giá trị cột trên từng dòng nằm ở `SharePointItem.fields`. Ẩn/hiện cột chỉ lọc UI. Query OData encode `%20` (không `+`).

## React — thứ tự đọc

```text
types/                      Props, Messages, MSAL
auth/                       createMsalTokenProvider
provider/                   AppProvider + SharePointProvider + hooks
hooks/                      useFolderChildren, useLibraryFields, getErrorMessage
fluent/                     theme, isDarkTheme
i18n/                       messages
utils/                      format bytes / ngày
components/file-manager/    UI duyệt thư viện (browser, bảng, empty, banner, icons)
```

## Quy tắc khi sửa

- Interface/type nằm ở `types/` — không khai báo chung file logic.
- Việc liên quan nằm **cùng thư mục**; mỗi file một việc.
- REST chỉ đọc: thêm ghi/xóa thì tách service mới, không nhồi vào `services/library`.
- `$select` chỉ khi lấy 1–vài property (vd. `ServerRelativeUrl`). List items / fields: **không** `$select` để đủ cột. `$expand` cho `File`, `Folder`, `Author`, `Editor`. Không `$skip` trên items — dùng `@odata.nextLink`.
- Comment theo **khối / hàm** (ý định + vì sao), tiếng Việt, không comment từng dòng gán biến.
