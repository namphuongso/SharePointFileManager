import { InteractionStatus } from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";
import {
  createMsalTokenProvider,
  defaultSharePointScopes,
  SharePointAppProvider,
  SharePointFileManager,
} from "@namphuongso/sharepoint-file-manager";
import { useMemo } from "react";
import "@namphuongso/sharepoint-file-manager/styles.css";

/**
 * Demo — chỉ 2 chỗ:
 * 1) SharePointAppProvider + config (siteUrl + token)
 * 2) SharePointFileManager libraryName
 */
export function App() {
  const { instance, accounts, inProgress } = useMsal();
  const account = accounts[0];
  const siteUrl = import.meta.env.VITE_SITE_URL as string | undefined;
  const libraryName = (import.meta.env.VITE_LIBRARY_NAME as string | undefined) || "Documents";
  const scopes = useMemo(
    () => (siteUrl ? defaultSharePointScopes(siteUrl) : []),
    [siteUrl],
  );

  const tokenProvider = useMemo(() => {
    if (!account) return undefined;
    return createMsalTokenProvider({ instance, account });
  }, [instance, account]);

  if (inProgress !== InteractionStatus.None) {
    return <p>Loading Microsoft session...</p>;
  }

  if (!account || !tokenProvider) {
    return (
      <main style={{ minHeight: "100vh", background: "#f5f5f5" }}>
        <h1>SharePoint File Manager</h1>
        <button
          type="button"
          onClick={() =>
            instance.loginRedirect({
              scopes: scopes.length > 0 ? scopes : ["https://tcsvn.sharepoint.com/AllSites.Write"],
            })
          }
        >
          Sign in with Microsoft
        </button>
      </main>
    );
  }

  if (!siteUrl) {
    return (
      <p style={{ padding: 24 }}>
        Set <code>VITE_SITE_URL</code> (and optional <code>VITE_LIBRARY_NAME</code>) in .env.local
      </p>
    );
  }

  return (
    <SharePointAppProvider
      locale="vi-VN"
      config={{
        siteUrl,
        scopes,
        tokenProvider,
      }}
    >
      <main style={{ minHeight: "100vh", background: "#f5f5f5", display: "flex", flexDirection: "column" }}>
        <div style={{ marginBottom: 12, padding: 12 }}>
          Signed in as {account.username}{" "}
          <button type="button" onClick={() => instance.logoutRedirect()}>
            Sign out
          </button>
          <span style={{ marginLeft: 12, color: "#666" }}>library: {libraryName}</span>
        </div>
        <div style={{ flex: 1, minHeight: 0 }}>
          <SharePointFileManager libraryName={libraryName} className="h-full" />
        </div>
      </main>
    </SharePointAppProvider>
  );
}
