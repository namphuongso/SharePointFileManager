# Microsoft Graph API map (v1.0)

Library không giả định Graph làm được mọi thứ SharePoint web UI làm được. Bảng dưới là những gì **đang gọi**.

Base: `https://graph.microsoft.com/v1.0`

## Drive / items

| Việc | Method | Path |
|---|---|---|
| Default drive | GET | `/sites/{siteId}/drive` |
| Get item | GET | `/drives/{driveId}/items/{itemId}` |
| List children | GET | `/drives/{driveId}/items/{itemId}/children` |
| Create folder | POST | `.../children` body `{ name, folder: {} }` |
| Rename / move | PATCH | `.../items/{itemId}` |
| Delete | DELETE | `.../items/{itemId}` (vào recycle bin SharePoint) |
| Copy | POST | `.../items/{itemId}/copy` (202 + Location monitor) |
| Download | GET item `downloadUrl` hoặc `.../content` |
| Simple upload < 4MB | PUT | `.../items/{parent}:/{name}:/content` |
| Large upload | POST | `.../createUploadSession` rồi PUT chunk 320KiB |
| Search | GET | `.../items/{folderId}/search(q='...')` |
| Preview | POST | `.../preview` |
| Versions | GET | `.../versions` |
| Restore version | POST | `.../versions/{id}/restoreVersion` |

Copy/move **không** download rồi upload lại.

`$expand=thumbnails` trên children **không** dùng — Graph không hỗ trợ cho SharePoint/OneDrive for Business.

## Sharing / permission

| Việc | Method | Path |
|---|---|---|
| Invite user/group | POST | `.../invite` roles `read` \| `write` |
| Create link | POST | `.../createLink` type `view`\|`edit`, scope `anonymous`\|`organization`\|`users` |
| List permissions | GET | `.../permissions` |
| Update roles | PATCH | `.../permissions/{id}` |
| Remove | DELETE | `.../permissions/{id}` |

### Hạn chế

- `inheritedFrom` có mặt → permission kế thừa, **không** hiện nút Remove.
- Non-owner có thể không thấy full ACL.
- `anonymous` / Anyone có thể bị tenant policy chặn → Graph 403; UI hiện lỗi, không giả định mọi tenant bật Anyone.
- Expiration chủ yếu áp dụng sharing link trên SharePoint/OneDrive for Business.
- Không phải mọi permission đều PATCH/DELETE được dù SharePoint UI có nút tương ứng.

## Recycle bin

Không có trong v1. List recycle bin Graph vẫn chủ yếu beta.

## Error mapping

| HTTP / Graph | UI |
|---|---|
| 401 `unauthenticated` | Cần đăng nhập lại |
| 403 `accessDenied` | Không có quyền |
| 404 `itemNotFound` | Không còn tồn tại |
| 409 `nameAlreadyExists` | Trùng tên |
| 429 | Retry `Retry-After` |
