import { describe, expect, it } from "vitest";
import { PermissionKind } from "../../types/permissions";
import { hasPermissionKind, toItemCapabilities } from "./parse-base-permissions";

describe("hasPermissionKind", () => {
  it("detects ViewListItems at bit 0 (SPBasePermissions value 1)", () => {
    expect(hasPermissionKind({ High: "0", Low: "1" }, PermissionKind.ViewListItems)).toBe(true);
    expect(hasPermissionKind({ High: "0", Low: "0" }, PermissionKind.ViewListItems)).toBe(false);
  });

  it("detects DeleteListItems at bit 3 (value 8)", () => {
    expect(hasPermissionKind({ High: "0", Low: "8" }, PermissionKind.DeleteListItems)).toBe(true);
    expect(hasPermissionKind({ High: "0", Low: "16" }, PermissionKind.DeleteListItems)).toBe(false);
  });

  it("detects OpenItems at bit 5 (value 32)", () => {
    expect(hasPermissionKind({ High: "0", Low: "32" }, PermissionKind.OpenItems)).toBe(true);
  });

  it("detects ManagePermissions at bit 25 in Low word (value 2^25)", () => {
    expect(
      hasPermissionKind({ High: "0", Low: String(1 << 25) }, PermissionKind.ManagePermissions),
    ).toBe(true);
  });

  it("decodes Contribute-like mask including DeleteListItems", () => {
    // View+Add+Edit+Delete+OpenItems = 1+2+4+8+32 = 47
    const raw = { High: "0", Low: "47" };
    expect(hasPermissionKind(raw, PermissionKind.ViewListItems)).toBe(true);
    expect(hasPermissionKind(raw, PermissionKind.AddListItems)).toBe(true);
    expect(hasPermissionKind(raw, PermissionKind.EditListItems)).toBe(true);
    expect(hasPermissionKind(raw, PermissionKind.DeleteListItems)).toBe(true);
    expect(hasPermissionKind(raw, PermissionKind.OpenItems)).toBe(true);
    expect(hasPermissionKind(raw, PermissionKind.ApproveItems)).toBe(false);
  });

  it("decodes screenshot mask (High 688, Low 1006834415)", () => {
    const raw = { High: "688", Low: "1006834415" };
    expect(hasPermissionKind(raw, PermissionKind.ViewListItems)).toBe(true);
    expect(hasPermissionKind(raw, PermissionKind.AddListItems)).toBe(true);
    expect(hasPermissionKind(raw, PermissionKind.DeleteListItems)).toBe(true);
    expect(hasPermissionKind(raw, PermissionKind.OpenItems)).toBe(true);
  });
});

describe("toItemCapabilities", () => {
  it("maps ViewListItems and AddListItems from SPBasePermissions flags", () => {
    // View=1 + Add=2 → Low 3
    const caps = toItemCapabilities({ High: "0", Low: "3" });
    expect(caps.canView).toBe(true);
    expect(caps.canAdd).toBe(true);
    expect(caps.canEdit).toBe(false);
    expect(caps.canDelete).toBe(false);
  });

  it("maps DeleteListItems for Contribute-like mask", () => {
    const caps = toItemCapabilities({ High: "0", Low: "47" });
    expect(caps.canDelete).toBe(true);
    expect(caps.canOpen).toBe(true);
  });
});
