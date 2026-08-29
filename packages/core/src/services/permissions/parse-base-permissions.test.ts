import { describe, expect, it } from "vitest";
import { PermissionKind } from "../../types/permissions";
import { hasPermissionKind, toItemCapabilities } from "./parse-base-permissions";

describe("hasPermissionKind", () => {
  it("detects ViewListItems at bit 1 (Microsoft PermissionKind.ViewListItems = 1)", () => {
    expect(hasPermissionKind({ High: "0", Low: "2" }, PermissionKind.ViewListItems)).toBe(true);
    expect(hasPermissionKind({ High: "0", Low: "0" }, PermissionKind.ViewListItems)).toBe(false);
  });

  it("detects OpenItems at bit 6", () => {
    expect(hasPermissionKind({ High: "0", Low: "64" }, PermissionKind.OpenItems)).toBe(true);
  });

  it("detects ManagePermissions at bit 26 in Low word", () => {
    expect(hasPermissionKind({ High: "0", Low: "67108864" }, PermissionKind.ManagePermissions)).toBe(
      true,
    );
  });

  it("decodes screenshot mask (High 688, Low 1006834415) as full access", () => {
    const raw = { High: "688", Low: "1006834415" };
    expect(hasPermissionKind(raw, PermissionKind.ViewListItems)).toBe(true);
    expect(hasPermissionKind(raw, PermissionKind.AddListItems)).toBe(true);
    expect(hasPermissionKind(raw, PermissionKind.ManagePermissions)).toBe(true);
  });
});

describe("toItemCapabilities", () => {
  it("maps ViewListItems and AddListItems", () => {
    const caps = toItemCapabilities({ High: "0", Low: "6" });
    expect(caps.canView).toBe(true);
    expect(caps.canAdd).toBe(true);
    expect(caps.canEdit).toBe(false);
    expect(caps.canDelete).toBe(false);
  });
});
