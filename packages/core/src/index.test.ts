import { describe, expect, it } from "vitest";
import { mapRestError, mapStatusToCode, resolveConfig, SharePointErrorCode } from "./index";

describe("resolveConfig", () => {
  it("requires siteUrl and defaults SharePoint scopes", () => {
    const config = resolveConfig({
      siteUrl: "https://contoso.sharepoint.com/sites/eOffice",
      tokenProvider: { getAccessToken: async () => "token" },
    });
    expect(config.siteId).toContain("contoso.sharepoint.com");
    expect(config.scopes[0]).toContain("AllSites.Write");
    expect(config.rootItemId).toBe("root");
  });

  it("strips trailing slash from siteUrl", () => {
    const config = resolveConfig({
      siteUrl: "https://contoso.sharepoint.com/sites/eOffice/",
      tokenProvider: { getAccessToken: async () => "token" },
    });
    expect(config.siteUrl).toBe("https://contoso.sharepoint.com/sites/eOffice");
  });
});

describe("mapRestError", () => {
  it("maps status codes", () => {
    expect(mapStatusToCode(403)).toBe(SharePointErrorCode.Forbidden);
    expect(
      mapRestError({ status: 403, body: { error: { message: { value: "Denied" } } } }).message,
    ).toBe("Denied");
  });
});
