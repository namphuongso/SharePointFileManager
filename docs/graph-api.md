# Microsoft Graph API map (v1.0)

Library không giả định Graph làm được mọi thứ SharePoint web UI làm được. Bảng dưới là những gì **đang gọi**.

Base: `https://graph.microsoft.com/v1.0`

Các hợp đồng được đối chiếu với tài liệu Microsoft Learn hiện hành: [DriveItem upload](https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0), [upload session](https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0), [copy](https://learn.microsoft.com/en-us/graph/api/driveitem-copy?view=graph-rest-1.0), [permissions](https://learn.microsoft.com/en-us/graph/api/permission-update?view=graph-rest-1.0), [activities](https://learn.microsoft.com/en-us/graph/api/itemactivity-list?view=graph-rest-1.0), [checkout](https://learn.microsoft.com/en-us/graph/api/driveitem-checkout?view=graph-rest-1.0), và [search files](https://learn.microsoft.com/en-us/graph/search-concept-files).

## Drive / items

| Việc | Method | Path |
|---|---|---|
| Default drive | GET | `/sites/{siteId}/drive` |
| List drives | GET | `/sites/{siteId}/drives` |
| Get item | GET | `/drives/{driveId}/items/{itemId}` |
| List children | GET | `/drives/{driveId}/items/{itemId}/children` |
| List library items (security-trimmed) | GET | `/sites/{siteId}/lists/{listId}/items?$expand=driveItem` |
| Create folder | POST | `.../children` body `{ name, folder: {} }` |
| Rename / move | PATCH | `.../items/{itemId}` |
| Delete | DELETE | `.../items/{itemId}` (vào recycle bin SharePoint) |
| Copy | POST | `.../items/{itemId}/copy` (202 + Location monitor) |
| Download | GET item `downloadUrl` hoặc `.../content` |
| Simple upload < 4MB | PUT | `.../items/{parent}:/{name}:/content` |
| Large upload | POST | `.../createUploadSession` rồi PUT chunk (bội 320 KiB, mặc định ~10 MiB) |
| Search (folder) | GET | `.../items/{folderId}/search(q='...')` |
| Search (library) | POST | `/search/query` (`entityTypes: driveItem`) |
| Preview | POST | `.../preview` |
| Versions | GET | `.../versions` |
| Restore version | POST | `.../versions/{id}/restoreVersion` |
| Checkout / checkin / discard | POST | `.../checkout`, `.../checkin`, `.../discardCheckout` |
| Activities | GET | `.../items/{itemId}/activities` |

Copy/move **không** download rồi upload lại.

`$expand=thumbnails` trên children đang được dùng để giảm số request preview. Hành vi có thể khác nhau theo tenant/workload nên cần fallback nếu thumbnail rỗng.

### Item-level share / security-trimmed library

Khi user chỉ có quyền trên file (không browse được `children` của thư mục gốc), Graph `.../root/children` trả rỗng. SharePoint web vẫn hiện file vì view library đọc **list items**.

Fallback chính thống:

`GET /sites/{siteId}/lists/{listId}/items?$expand=driveItem`

API này security-trim theo user đang login — cùng nguồn với list view SharePoint. File trong subfolder vẫn hiện; nếu user thấy được thư mục chứa chúng thì UI giữ cây thư mục, không dàn phẳng lung tung.

## Sharing / permission

| Việc | Method | Path |
|---|---|---|
| Invite user/group | POST | `.../invite` roles `read` \| `write` |
| Create link | POST | `.../createLink` type `view`\|`edit`, scope `anonymous`\|`organization` (specific people → `invite`) |
| List permissions | GET | `.../permissions` |
| Update roles | PATCH | `.../permissions/{id}` |
| Remove | DELETE | `.../permissions/{id}` |
| People picker | GET | `/me/people?$search=` |
| Directory users | GET | `/users?$search=` (`ConsistencyLevel: eventual`) |
| Directory groups | GET | `/groups?$filter=startswith(displayName,'...')` |

### Hạn chế

- `inheritedFrom` có mặt → permission kế thừa, **không** hiện nút Remove.
- Non-owner có thể không thấy full ACL.
- `anonymous` / Anyone có thể bị tenant policy chặn → Graph 403; UI hiện lỗi, không giả định mọi tenant bật Anyone.
- Expiration chủ yếu áp dụng sharing link trên SharePoint/OneDrive for Business.
- Không phải mọi permission đều PATCH/DELETE được dù SharePoint UI có nút tương ứng.

## Error mapping

| HTTP / Graph | UI |
|---|---|
| 401 `unauthenticated` | Cần đăng nhập lại |
| 403 `accessDenied` | Không có quyền |
| 404 `itemNotFound` | Không còn tồn tại |
| 409 `nameAlreadyExists` | Trùng tên |
| 429 | Retry `Retry-After` |
