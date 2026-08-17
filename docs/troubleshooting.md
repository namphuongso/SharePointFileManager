# Troubleshooting

## Login lần hai / popup MSAL

Library không được tự login. Chỉ truyền `createMsalTokenProvider({ instance, account })` từ **cùng** `PublicClientApplication` của app.

## 401 liên tục

- Account chưa login
- Scopes Graph chưa consent
- App registration chưa có Graph delegated permissions

Xin scopes lúc login hoặc admin consent.

## 403 khi share / Anyone

Tenant có thể tắt anonymous link. Dùng `organization` hoặc `users`. Đây không phải bug library.

## 403 khi Remove access

Permission đang inherited. UI phải hiện "Kế thừa từ thư mục cha" và không gọi DELETE.

## Token backend không gọi được Graph

Đúng. Access token của API nội bộ không dùng cho Graph. Phải `acquireTokenSilent` với Graph scopes.

## Preview trống

Graph preview chủ yếu Office / PDF / ảnh. File khác: Open in SharePoint.

## Copy treo

Copy Graph là async (202). Nếu monitor URL fail, kiểm tra CORS/auth của Location header.

## Tailwind đụng class host

Library dùng prefix `spm-` và CSS đã compile. Host chỉ import `styles.css`.
