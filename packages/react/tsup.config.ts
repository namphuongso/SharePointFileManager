import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm", "cjs"],
  dts: true,
  sourcemap: true,
  clean: true,
  treeshake: true,
  splitting: false,
  external: [
    "react",
    "react-dom",
    "@tanstack/react-query",
    "@azure/msal-browser",
    "@fluentui/react-components",
    "@fluentui/react-icons",
    "@namphuongso/sharepoint-file-manager-core",
  ],
});
