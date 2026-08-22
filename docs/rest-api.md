# SharePoint REST API map

Library chỉ **GET** các endpoint sau:

| Việc | Path |
|---|---|
| Web | `/_api/web` |
| Resolve library | `/_api/web/lists/getbytitle('…')` + `$expand=RootFolder` |
| Folders / files | `/_api/web/GetFolderByServerRelativeUrl('…')/Folders` và `/Files` |
| Get by UniqueId | `/_api/web/GetFolderById('…')` |

Auth: `defaultSharePointScopes(siteUrl)` → `https://{host}/AllSites.Write`.
