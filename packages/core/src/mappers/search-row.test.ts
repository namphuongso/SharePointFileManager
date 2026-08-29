import { describe, expect, it } from "vitest";
import { mapSearchRow } from "./search-row";

const libraryUrl = "https://contoso.sharepoint.com/NPPDocumentDev";

function row(cells: Record<string, string>) {
  return {
    Cells: Object.entries(cells).map(([Key, Value]) => ({ Key, Value })),
  };
}

describe("mapSearchRow", () => {
  it("maps file with UniqueId braces stripped", () => {
    const item = mapSearchRow(
      row({
        Title: "123.txt",
        Filename: "123.txt",
        Path: `${libraryUrl}/Thư mục/123.txt`,
        UniqueId: "{aaef1547-b160-44dd-99b3-bfa7cd043dbc}",
        IsContainer: "false",
        Size: "42",
        LastModifiedTime: "2026-08-24T00:00:00Z",
      }),
      libraryUrl,
    );
    expect(item).toEqual({
      id: "aaef1547-b160-44dd-99b3-bfa7cd043dbc",
      name: "123.txt",
      type: "file",
      size: 42,
      lastModifiedDateTime: "2026-08-24T00:00:00Z",
    });
  });

  it("maps folder when IsContainer true", () => {
    const item = mapSearchRow(
      row({
        Filename: "1222",
        Path: `${libraryUrl}/Thư mục/1222`,
        UniqueId: "bbef1547-b160-44dd-99b3-bfa7cd043dbc",
        IsContainer: "true",
      }),
      libraryUrl,
    );
    expect(item?.type).toBe("folder");
    expect(item?.size).toBeUndefined();
  });

  it("skips library root and Forms paths", () => {
    expect(
      mapSearchRow(
        row({
          Filename: "NPPDocumentDev",
          Path: libraryUrl,
          UniqueId: "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
          IsContainer: "true",
        }),
        libraryUrl,
      ),
    ).toBeUndefined();
    expect(
      mapSearchRow(
        row({
          Filename: "AllItems.aspx",
          Path: `${libraryUrl}/Forms/AllItems.aspx`,
          UniqueId: "cccccccc-cccc-cccc-cccc-cccccccccccc",
          IsContainer: "false",
        }),
        libraryUrl,
      ),
    ).toBeUndefined();
  });

  it("maps Author/Editor into fields for column UI", () => {
    const item = mapSearchRow(
      row({
        Filename: "a.txt",
        Path: `${libraryUrl}/a.txt`,
        UniqueId: "aaef1547-b160-44dd-99b3-bfa7cd043dbc",
        IsContainer: "false",
        Author: "Quoc Van Ly",
        ModifiedBy: "Service",
      }),
      libraryUrl,
      ["Author", "Editor"],
    );
    expect(item?.fields?.Author).toEqual({ Title: "Quoc Van Ly" });
    expect(item?.fields?.Editor).toEqual({ Title: "Service" });
  });

  it("skips row without UniqueId", () => {
    expect(
      mapSearchRow(
        row({
          Filename: "x.txt",
          Path: `${libraryUrl}/x.txt`,
          IsContainer: "false",
        }),
        libraryUrl,
      ),
    ).toBeUndefined();
  });
});
