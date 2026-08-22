# Troubleshooting

## Login / popup MSAL

Library không tự login. Host truyền `tokenProvider` từ cùng `PublicClientApplication`.

## 401

- Chưa login
- Scope `AllSites.Write` chưa consent
- App registration thiếu SharePoint delegated permission

## Token backend không gọi được SharePoint

Access token API nội bộ không dùng cho SharePoint REST. Phải `acquireTokenSilent` với SharePoint scopes.
