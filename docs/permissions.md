# Microsoft Entra delegated permissions

Library dùng **delegated** token của user đang login. Không dùng application permissions, không dùng client secret.

## Default scopes

```text
Files.ReadWrite
Sites.ReadWrite.All
User.Read.All
People.Read
Directory.Read.All
```

Đủ browse + upload + share + people picker trên document library. Browse khi `children` rỗng dùng list-item `$expand=driveItem` (không cần `Files.Read.All` / `sharedWithMe`).

## Least privilege theo feature

| Feature | Quyền tối thiểu (tham khảo Graph docs) |
|---|---|
| List / download / search / preview / versions | `Files.ReadWrite` hoặc `Sites.Read.All` / `Sites.ReadWrite.All` trên SharePoint site |
| Security-trimmed library listing (`lists/.../items?$expand=driveItem`) | `Sites.Read.All` / `Sites.ReadWrite.All` |
| Upload / rename / delete / copy / move / folder | `Files.ReadWrite` / `Sites.ReadWrite.All` |
| Invite / createLink / remove permission | `Files.ReadWrite` ; với library SharePoint thường cần `Sites.ReadWrite.All` |
| People picker (tên người / nhóm trong tenant) | `People.Read` + `User.Read.All` ; nhóm cần `Directory.Read.All` |

`Sites.Selected` có thể hẹp hơn nhưng phải được admin gán site; không mặc định.

Không cấp Application `Sites.ReadWrite.All` cho SPA.

## App registration

1. Entra ID → App registrations → app **của project** (đã là SPA)
2. Authentication: SPA redirect URI đã có
3. API permissions → Microsoft Graph → Delegated → thêm quyền trên
4. Grant admin consent
5. Đưa scopes vào MSAL `loginRequest`

Mỗi project dùng `clientId` / `siteId` / `driveId` của project đó.

## Permission safety

Microsoft Graph không trả về `inheritedFrom` ổn định cho OneDrive for Business/SharePoint document libraries. Library chỉ hiển thị thao tác Remove cho sharing link trực tiếp; permission user/group được coi là không xác định để tránh xóa nhầm ACL kế thừa.
