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
client.ts                   Ghép config + REST + cache thư viện + FolderService + FieldService + PermissionService
auth/                       defaultSharePointScopes
config/                     resolve-config, create-sharepoint-config
rest/                       client GET, build-url, parse-body, throttle, odata
services/library/           resolve theo libraryName
services/folder/            listChildren + resolve path
services/fields/            defaultView/viewfields + GET /fields; fallback GET /fields
services/permissions/       EffectiveBasePermissions (UniqueId) → ItemCapabilities
services/search/            GET search/query (Path thư viện, security trim) → flat list
mappers/                    list item / SP.Field / search row → model
errors/                     SharePointError + map HTTP/OData
```

Luồng list một folder:

1. `getLibrary()` (cache) → `listId` + `rootFolderServerRelativeUrl`
2. Nếu không phải root: `GetFolderById` lấy `ServerRelativeUrl` (= FileDirRef)
3. Cột option: `GET web/lists(guid)/defaultView/viewfields`, sau đó lấy `Title` của đúng các `InternalName` từ `/fields` theo locale; view lỗi → `GET /fields` chỉ InternalName 3 cột cố định (`FileLeafRef`, `Modified`, `File_x0020_Size`) — không lấy hết schema
4. `GET web/lists(guid)/items?$filter=FileDirRef eq '...'` — `$select` các cột bước 3 trừ Name/Modified/Size (render từ File/Folder) và computed fields + File/Folder; `$expand=File,Folder` và thêm `Author`/`Editor` chỉ khi cột đó nằm trong `$select` (`Author/Id,Author/Title`), `$top=30`, `$orderby=FSObjType desc,FileLeafRef` (mặc định). Click cột: GET lại trang đầu với `$orderby=FSObjType desc,{InternalName} asc|desc` (`Author/Title` / `Editor/Title`). Không sort client khi còn `@odata.nextLink`.
5. Trang sau: GET nguyên `@odata.nextLink` (`SharePointRestClient.getUrl`)
6. Bỏ folder `Forms`, map `File.UniqueId` / `Folder.UniqueId`; folder: `childItemCount` = `Folder.ItemCount` (tổng con trực tiếp)
7. Cột bị báo `does not exist` (ghost theo tenant): `FieldService.exclude` loại khỏi danh sách rồi thử lại mỗi GET — tối đa 3 lần

Ẩn/hiện cột chỉ trên UI (`ColumnPicker`): mặc định hiện Tên / Sửa đổi / Kích thước + Người sửa đổi (`Editor`) giống All Documents; tick thêm cột không gọi REST. Cột `ItemChildCount` trên view (nếu bật) hiện `childItemCount`; không dùng `FolderChildCount`. Không `storagemetrics` / `getMetrics`.

Nhãn cột = Title SharePoint trả theo locale (`Accept-Language`); `messages.fieldLabels` chỉ để host ghi đè. Đổi locale runtime qua `SharePointClient.setLocale`; React Query cache tách theo locale.
Toolbar có menu VI/EN mặc định (`showLanguageSwitcher={false}` để ẩn).

Giá trị cột trên từng dòng nằm ở `SharePointItem.fields`. Query OData encode `%20` (không `+`).

## Quyền (PermissionService)

Chỉ **GET** `EffectiveBasePermissions` — quyền user hiện tại, bitmask Microsoft `{ High, Low }` → `ItemCapabilities` (`canAdd`, `canEdit`, …).

```text
SharePointClient.permissions
  getLibraryCapabilities()              listId — gate coarse khi mở thư viện
  getFolderCapabilities(uniqueId)       GetItemByUniqueId + effectiveBasePermissions; "root" → rootFolderUniqueId
  getFileCapabilities(uniqueId)         GetItemByUniqueId + effectiveBasePermissions
  getItemCapabilities(type, uniqueId)   wrapper file | folder
```

Không GET quyền từng dòng `listChildren`. **React:** `useFolderViewCapabilities(folderUniqueId)` gate `canView` trước `useFolderChildren`; không quyền → `ForbiddenState`, không gọi list items. Action mới: gate UI bằng capability tương ứng — xem `.cursor/rules/permissions.mdc`.

## Search — item được xem trong thư viện

Tab **Có quyền xem** dùng `GET /_api/search/query` (SharePoint Search REST, không Graph). Security trim theo token; giới hạn KQL `Path` thư viện hiện tại. Flat list + `StartRow` + `sortlist` + cột option (managed properties) — không thay browse folder.

```text
SharePointClient.search
  listAccessible({ startRow, rowLimit, sort, fieldInternalNames })
```

**React:** cùng ColumnPicker / sort header / infinite scroll như Home; mở folder → chuyển Home + breadcrumb UniqueId.

## React — thứ tự đọc

```text
types/                      Props, Messages, MSAL
auth/                       createMsalTokenProvider
provider/                   AppProvider + SharePointProvider + hooks
hooks/                      useFolderChildren, useAccessibleItems, useFolderViewCapabilities, useLibraryFields, useColumnSort, getErrorMessage
fluent/                     theme, isDarkTheme
i18n/                       messages
utils/                      format bytes / ngày
components/file-manager/    UI duyệt thư viện (browser, bảng, empty, banner, icons)
```

## Quy tắc khi sửa

- Interface/type nằm ở `types/` — không khai báo chung file logic.
- Việc liên quan nằm **cùng thư mục**; mỗi file một việc.
- REST chỉ đọc: thêm ghi/xóa thì tách service mới, không nhồi vào `services/library`.
- `$select` chỉ property cần: `GetFolderById` → `ServerRelativeUrl`; list items → `listItemSelect(internalNames)` (không `*`); `/fields` → `Id,Title,InternalName`. `$expand` cho `File`, `Folder`, `Author`, `Editor`. Không `$skip` trên items — dùng `@odata.nextLink`.
- Comment theo **khối / hàm** (ý định + vì sao), tiếng Việt, không comment từng dòng gán biến.
