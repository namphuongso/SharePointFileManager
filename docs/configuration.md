# Configuration

```ts
interface SharePointConfig {
  siteId: string;
  driveId?: string;       // nếu bỏ trống, lấy default drive của site
  rootItemId?: string;    // mặc định "root"
  scopes?: string[];      // mặc định Files.ReadWrite + Sites.ReadWrite.All
  graphBaseUrl?: string;  // mặc định https://graph.microsoft.com/v1.0
  tokenProvider: TokenProvider;
  features?: FeatureConfig;
}
```

`tenantId` / `clientId` không bắt buộc khi nhúng app thật — MSAL host đã có.

## Multi-project

```ts
// Project A
{ siteId: "site-a", driveId: "drive-a", tokenProvider }

// Project B
{ siteId: "site-b", driveId: "drive-b", tokenProvider }
```

Library không biết nghiệp vụ project.

## Features

Mọi flag mặc định `true` trừ recycle bin (không có trong v1).

Action trên UI = feature flag AND kết quả Graph. 403 → permission denied.

## Locale / theme

```tsx
<SharePointFileManager locale="vi-VN" config={config} />
```

CSS variables:

```css
.spm-root {
  --spm-primary: #2563eb;
  --spm-danger: #dc2626;
}
```
