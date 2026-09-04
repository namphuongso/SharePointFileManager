import { describe, expect, it } from "vitest";
import { addFolderUsingPath } from "./add-folder-using-path";

describe("addFolderUsingPath", () => {
  it("keeps slashes in server-relative path", () => {
    expect(addFolderUsingPath("/NPPDocumentDev/Reports")).toBe(
      "web/Folders/AddUsingPath(decodedurl='/NPPDocumentDev/Reports')",
    );
  });

  it("encodes spaces in segments as %20", () => {
    expect(addFolderUsingPath("/Lib/My Folder")).toBe(
      "web/Folders/AddUsingPath(decodedurl='/Lib/My%20Folder')",
    );
  });
});
