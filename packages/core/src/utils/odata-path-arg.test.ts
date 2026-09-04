import { describe, expect, it } from "vitest";
import { encodeODataQuotedValue, encodeServerRelativePathArg } from "./odata-path-arg";

describe("encodeServerRelativePathArg", () => {
  it("preserves path separators", () => {
    expect(encodeServerRelativePathArg("/a/b/c")).toBe("/a/b/c");
  });

  it("encodes spaces per segment", () => {
    expect(encodeServerRelativePathArg("/Shared Documents/My Folder")).toBe(
      "/Shared%20Documents/My%20Folder",
    );
  });
});

describe("encodeODataQuotedValue", () => {
  it("encodes leaf names", () => {
    expect(encodeODataQuotedValue("a b.txt")).toBe("a%20b.txt");
  });
});
