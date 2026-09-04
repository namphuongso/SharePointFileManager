import { describe, expect, it } from "vitest";
import { listItemRecyclePath } from "./delete-item-path";

describe("listItemRecyclePath", () => {
  it("builds GetItemByUniqueId recycle path", () => {
    expect(
      listItemRecyclePath(
        "11111111-1111-1111-1111-111111111111",
        "aaef1547-b160-44dd-99b3-bfa7cd043dbc",
      ),
    ).toBe(
      "web/lists(guid'11111111-1111-1111-1111-111111111111')/GetItemByUniqueId(guid'aaef1547-b160-44dd-99b3-bfa7cd043dbc')/recycle()",
    );
  });

  it("strips braces from GUIDs", () => {
    expect(
      listItemRecyclePath(
        "{11111111-1111-1111-1111-111111111111}",
        "{aaef1547-b160-44dd-99b3-bfa7cd043dbc}",
      ),
    ).toBe(
      "web/lists(guid'11111111-1111-1111-1111-111111111111')/GetItemByUniqueId(guid'aaef1547-b160-44dd-99b3-bfa7cd043dbc')/recycle()",
    );
  });
});
