# Architecture

## Nguyên tắc

- SharePoint Online là nguồn dữ liệu và nguồn sự thật về quyền.
- Microsoft Entra ID là identity provider.
- Microsoft Graph là API layer.
- Library không lưu ACL, không mirror permission, không chứa client secret.

```text
Application (MSAL already signed in)
    │ TokenProvider / createMsalTokenProvider
    ▼
@namphuongso/sharepoint-file-manager   React UI, hooks, TanStack Query
    │ bundles
    ▼
@namphuongso/sharepoint-file-manager-core   GraphClient + services
    ▼
Microsoft Graph v1.0
    ▼
SharePoint Online
```

## Packages

| Package | Publish | Vai trò |
|---|---|---|
| `packages/core` | private workspace | TokenProvider, Graph client, domain model, services |
| `packages/react` | `@namphuongso/sharepoint-file-manager` | UI + hooks; giữ core là dependency external để bundle browser không kéo nhánh Node vào client |

React components không `fetch` Graph. Mọi request đi qua `GraphClient`.

## Auth

Production chỉ có:

```ts
createMsalTokenProvider({ instance, account })
```

Library không tạo `PublicClientApplication`, không `loginPopup` / `loginRedirect`.

## State

TanStack Query cache folder children (paged + infinite), accessible library items, search, permissions, versions. Mutations (upload, rename, delete, copy, move, checkout, restore) invalidate `children`, `children-infinite`, and `accessible-library-items` for the affected folder/library.

## Embed model

```text
SharePointAppProvider (siteId + token + features)     ← config một lần ở app
    └── SharePointFileManager libraryName="…"         ← từng module/route
            └── embedded UI = command bar + file list  ← không header / shell nav
```

Standalone demo có thể truyền `config` đầy đủ + `embedded={false}` để hiện full `SharePointShell`.

## UI / Tailwind

Class prefix `spm-`. Host import `styles.css` đã compile. Không bắt host scan source Tailwind của library.

## Ngoài v1

- Theme package riêng
- Library tự MSAL login
- Retention
