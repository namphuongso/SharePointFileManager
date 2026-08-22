import type { SharePointLibraryTarget } from "@namphuongso/sharepoint-file-manager-core";
import { SharePointProvider } from "../../provider/SharePointProvider";
import { useOptionalSharePointApp } from "../../provider/useSharePointApp";
import type { SharePointFileManagerProps } from "../../types";
import { FileBrowser } from "./FileBrowser";

/**
 * Điểm vào UI: ghép config (prop hoặc AppProvider) rồi bọc FileBrowser.
 * Host phải bọc SharePointAppProvider hoặc truyền `config`.
 */
export function SharePointFileManager(props: SharePointFileManagerProps) {
  const app = useOptionalSharePointApp();
  const target: SharePointLibraryTarget = {
    libraryName: props.libraryName,
    listId: props.listId,
    rootItemId: props.rootItemId,
  };
  const hasTarget = Boolean(target.libraryName || target.listId || target.rootItemId);

  if (!props.config && app && app.status === "error") {
    return (
      <p style={{ padding: 16, color: "#616161", fontSize: 14 }}>
        Không tải được cấu hình SharePoint. Kiểm tra `siteUrl` và quyền SharePoint.
      </p>
    );
  }

  const config =
    props.config ?? (app?.status === "ready" && app.appConfig ? app.createConfig(target) : undefined);

  if (!config) {
    throw new Error(
      "SharePointFileManager requires `config`, or `libraryName` inside SharePointAppProvider",
    );
  }

  return (
    <SharePointProvider
      config={hasTarget && props.config ? { ...props.config, ...target } : config}
      locale={props.locale ?? app?.locale}
      messages={props.messages ?? app?.messages}
      theme={props.theme}
    >
      <FileBrowser className={props.className} title={props.title} />
    </SharePointProvider>
  );
}
