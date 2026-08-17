# Microsoft Entra delegated permissions

Library dùng **delegated** token của user đang login. Không dùng application permissions, không dùng client secret.

## Default scopes

```text
Files.ReadWrite
Sites.ReadWrite.All
```

Đủ browse + upload + share trên document library.

## Least privilege theo feature

| Feature | Quyền tối thiểu (tham khảo Graph docs) |
|---|---|
| List / download / search / preview / versions | `Files.Read` + thường cần `Sites.Read.All` với SharePoint site |
| Upload / rename / delete / copy / move / folder | `Files.ReadWrite` |
| Invite / createLink / remove permission | `Files.ReadWrite` ; với library SharePoint thường cần `Sites.ReadWrite.All` |

`Sites.Selected` có thể hẹp hơn nhưng phải được admin gán site; không mặc định.

Không cấp Application `Sites.ReadWrite.All` cho SPA.

## App registration

1. Entra ID → App registrations → app **của project** (đã là SPA)
2. Authentication: SPA redirect URI đã có
3. API permissions → Microsoft Graph → Delegated → thêm quyền trên
4. Grant admin consent
5. Đưa scopes vào MSAL `loginRequest`

Mỗi project dùng `clientId` / `siteId` / `driveId` của project đó.
