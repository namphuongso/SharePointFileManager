import { describe, expect, it } from "vitest";
import {
  folderPermissionsPath,
  libraryPermissionsPath,
  listItemPermissionsPath,
} from "./effective-permissions-path";

describe("effective-permissions-path", () => {
  it("builds library path", () => {
    expect(libraryPermissionsPath("abc-def")).toBe(
      "web/lists(guid'abc-def')/effectiveBasePermissions",
    );
  });

  it("builds list item path with guid prefix", () => {
    expect(
      listItemPermissionsPath(
        "11111111-1111-1111-1111-111111111111",
        "aaef1547-b160-44dd-99b3-bfa7cd043dbc",
      ),
    ).toBe(
      "web/lists(guid'11111111-1111-1111-1111-111111111111')/GetItemByUniqueId(guid'aaef1547-b160-44dd-99b3-bfa7cd043dbc')/effectiveBasePermissions",
    );
  });

  it("strips braces from guids", () => {
    expect(
      listItemPermissionsPath(
        "{11111111-1111-1111-1111-111111111111}",
        "{aaef1547-b160-44dd-99b3-bfa7cd043dbc}",
      ),
    ).toBe(
      "web/lists(guid'11111111-1111-1111-1111-111111111111')/GetItemByUniqueId(guid'aaef1547-b160-44dd-99b3-bfa7cd043dbc')/effectiveBasePermissions",
    );
  });

  it("uses library path for root alias", () => {
    expect(
      folderPermissionsPath(
        "11111111-1111-1111-1111-111111111111",
        "root",
        "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
      ),
    ).toBe("web/lists(guid'11111111-1111-1111-1111-111111111111')/effectiveBasePermissions");
  });

  it("uses library path when uniqueId is rootFolderUniqueId", () => {
    expect(
      folderPermissionsPath(
        "11111111-1111-1111-1111-111111111111",
        "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
        "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
      ),
    ).toBe("web/lists(guid'11111111-1111-1111-1111-111111111111')/effectiveBasePermissions");
  });

  it("uses list item path for subfolder", () => {
    expect(
      folderPermissionsPath(
        "11111111-1111-1111-1111-111111111111",
        "bbef1547-b160-44dd-99b3-bfa7cd043dbc",
        "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
      ),
    ).toBe(
      "web/lists(guid'11111111-1111-1111-1111-111111111111')/GetItemByUniqueId(guid'bbef1547-b160-44dd-99b3-bfa7cd043dbc')/effectiveBasePermissions",
    );
  });
});
