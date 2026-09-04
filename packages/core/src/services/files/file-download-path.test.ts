import { describe, expect, it } from "vitest";
import { fileDownloadPath } from "./file-download-path";

describe("fileDownloadPath", () => {
  it("builds GetFileById /$value path", () => {
    expect(fileDownloadPath("aaef1547-b160-44dd-99b3-bfa7cd043dbc")).toBe(
      "web/GetFileById('aaef1547-b160-44dd-99b3-bfa7cd043dbc')/$value",
    );
  });
});
