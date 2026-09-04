import { describe, expect, it } from "vitest";
import { relativeFolderSegments } from "./relative-folder-segments";

describe("relativeFolderSegments", () => {
  it("returns empty at library root", () => {
    expect(relativeFolderSegments("/sites/a/Shared Documents", "/sites/a/Shared Documents")).toEqual(
      [],
    );
    expect(
      relativeFolderSegments("/sites/a/Shared Documents/", "/sites/a/Shared Documents"),
    ).toEqual([]);
  });

  it("splits nested folders", () => {
    expect(
      relativeFolderSegments(
        "/sites/a/Shared Documents",
        "/sites/a/Shared Documents/Reports/2024",
      ),
    ).toEqual(["Reports", "2024"]);
  });

  it("returns null when outside library", () => {
    expect(
      relativeFolderSegments("/sites/a/Shared Documents", "/sites/a/OtherLib/x"),
    ).toBeNull();
  });
});
