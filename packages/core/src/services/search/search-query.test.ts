import { describe, expect, it } from "vitest";
import { libraryAbsoluteUrl, searchQueryPath } from "./search-query-path";
import { librarySearchQuery } from "./search-query";

describe("search-query-path", () => {
  it("builds search/query path", () => {
    expect(searchQueryPath()).toBe("search/query");
  });

  it("joins siteUrl and root folder without double slash", () => {
    expect(libraryAbsoluteUrl("https://contoso.sharepoint.com/", "/NPPDocumentDev")).toBe(
      "https://contoso.sharepoint.com/NPPDocumentDev",
    );
  });
});

describe("librarySearchQuery", () => {
  it("scopes Path to library and excludes Forms", () => {
    const q = librarySearchQuery({
      siteUrl: "https://contoso.sharepoint.com",
      rootFolderServerRelativeUrl: "/NPPDocumentDev",
      rowLimit: 30,
      startRow: 0,
    });
    expect(q.querytext).toBe(
      `'Path:"https://contoso.sharepoint.com/NPPDocumentDev*" -Path:"https://contoso.sharepoint.com/NPPDocumentDev/Forms*"'`,
    );
    expect(q.selectproperties).toMatch(/^'/);
    expect(q.rowlimit).toBe(30);
    expect(q.startrow).toBe(0);
    expect(q.trimduplicates).toBe("false");
    expect(q.sortlist).toBeUndefined();
  });

  it("adds sortlist and extra managed properties", () => {
    const q = librarySearchQuery({
      siteUrl: "https://contoso.sharepoint.com",
      rootFolderServerRelativeUrl: "/NPPDocumentDev",
      rowLimit: 30,
      startRow: 30,
      fieldInternalNames: ["Editor", "Description"],
      sort: { field: "Modified", direction: "desc" },
    });
    expect(q.startrow).toBe(30);
    expect(q.sortlist).toBe("'LastModifiedTime:descending'");
    expect(String(q.selectproperties)).toContain("ModifiedBy");
    expect(String(q.selectproperties)).toContain("Description");
  });
});
