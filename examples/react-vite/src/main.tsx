import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import React from "react";
import ReactDOM from "react-dom/client";
import { App } from "./App";

const clientId = import.meta.env.VITE_CLIENT_ID as string | undefined;
const tenantId = import.meta.env.VITE_TENANT_ID as string | undefined;
const redirectUri = (import.meta.env.VITE_REDIRECT_URI as string | undefined) ?? window.location.origin;

const msal = new PublicClientApplication({
  auth: {
    clientId: clientId ?? "missing-client-id",
    authority: `https://login.microsoftonline.com/${tenantId ?? "common"}`,
    redirectUri,
  },
});

await msal.initialize();

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <MsalProvider instance={msal}>
      <App />
    </MsalProvider>
  </React.StrictMode>,
);
