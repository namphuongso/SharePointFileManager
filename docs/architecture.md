# Architecture

- SharePoint Online là nguồn dữ liệu.
- Entra ID là identity.
- API: **SharePoint REST** `GET /_api` (chỉ đọc danh sách). Không dùng Microsoft Graph.

```text
Application (MSAL already signed in)
    │ TokenProvider
    ▼
@namphuongso/sharepoint-file-manager   list UI
    ▼
@namphuongso/sharepoint-file-manager-core   REST GET + FolderService.listChildren
    ▼
SharePoint Online
```

## Packages

| Package | Vai trò |
|---|---|
| `packages/core` | Token, REST GET client, resolve library, list folders/files |
| `packages/react` | Bảng danh sách + breadcrumb |

UI không `fetch` trực tiếp. Mọi request đi qua `SharePointRestClient.get`.

Hiện **không** có upload, share, checkout, search, permissions UI.
