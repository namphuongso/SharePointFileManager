# Configuration

| Place | What |
|---|---|
| **`SharePointAppProvider`** | `siteUrl`, `tokenProvider`, optional scopes |
| **`SharePointFileManager`** | `libraryName` (or `listId`) |

```ts
interface SharePointAppConfig {
  siteId?: string;
  siteUrl?: string;
  scopes?: string[];
  tokenProvider: TokenProvider;
}

interface SharePointLibraryTarget {
  libraryName?: string;
  listId?: string;
  rootItemId?: string;
}
```

`siteUrl` bắt buộc. `siteId` chỉ dùng làm cache key (mặc định = `siteUrl`).
