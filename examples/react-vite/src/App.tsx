import { InteractionStatus } from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";
import { createMsalTokenProvider, SharePointFileManager } from "@namphuongso/sharepoint-file-manager";
import { useMemo } from "react";
import "@namphuongso/sharepoint-file-manager/styles.css";

const graphScopes = ["Files.ReadWrite", "Sites.ReadWrite.All", "User.Read.All", "People.Read", "Directory.Read.All"];

export function App() {
  const { instance, accounts, inProgress } = useMsal();
  const account = accounts[0];
  const siteId = import.meta.env.VITE_SITE_ID as string | undefined;
  const driveId = import.meta.env.VITE_DRIVE_ID as string | undefined;

  if (inProgress !== InteractionStatus.None) {
    return <p>Loading Microsoft session...</p>;
  }

  if (!account) {
    return (
      <main style={{ minHeight: "100vh", background: "#f5f5f5" }}>
        <h1>SharePoint File Manager</h1>
        <button
          type="button"
          onClick={() => instance.loginRedirect({ scopes: graphScopes })}
        >
          Sign in with Microsoft
        </button>
      </main>
    );
  }

  if (!siteId) {
    return <p style={{ padding: 24 }}>Set VITE_SITE_ID (and optional VITE_DRIVE_ID) in .env.local</p>;
  }

  const tokenProvider = useMemo(
    () => createMsalTokenProvider({ instance, account }),
    [instance, account],
  );

  const config = useMemo(
    () => ({ siteId, driveId, tokenProvider }),
    [siteId, driveId, tokenProvider],
  );

  return (
    <main style={{ minHeight: "100vh", background: "#f5f5f5" }}>
      <div style={{ marginBottom: 12 }}>
        Signed in as {account.username}{" "}
        <button type="button" onClick={() => instance.logoutRedirect()}>
          Sign out
        </button>
      </div>
      <SharePointFileManager locale="vi-VN" config={config} />
    </main>
  );
}
