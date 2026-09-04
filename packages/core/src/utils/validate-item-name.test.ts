import { describe, expect, it } from "vitest";
import { SharePointErrorCode } from "../errors/sharepoint-error";
import { assertValidItemName } from "./validate-item-name";

describe("assertValidItemName", () => {
  it("accepts normal names", () => {
    expect(() => assertValidItemName("Reports", "folder")).not.toThrow();
    expect(() => assertValidItemName("a.txt", "file")).not.toThrow();
  });

  it("rejects empty and invalid characters", () => {
    expect(() => assertValidItemName("  ", "folder")).toThrow();
    expect(() => assertValidItemName("a/b", "file")).toThrow();
    try {
      assertValidItemName("x:y", "folder");
    } catch (error) {
      expect((error as { code: string }).code).toBe(SharePointErrorCode.Unsupported);
    }
  });

  it("rejects reserved device names", () => {
    expect(() => assertValidItemName("CON", "folder")).toThrow();
    expect(() => assertValidItemName("com1.txt", "file")).toThrow();
  });
});
