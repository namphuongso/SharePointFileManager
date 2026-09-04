import { describe, expect, it } from "vitest";
import { uploadFileByFolderPath } from "./upload-file-by-folder-path";

describe("uploadFileByFolderPath", () => {
  it("builds GetFolderByServerRelativePath Files/add", () => {
    expect(uploadFileByFolderPath("/NPPDocumentDev", "a.txt", false)).toBe(
      "web/GetFolderByServerRelativePath(decodedUrl='/NPPDocumentDev')/Files/add(url='a.txt',overwrite=false)",
    );
  });

  it("encodes folder spaces and file name", () => {
    expect(uploadFileByFolderPath("/Lib/My Folder", "my file.pdf", true)).toBe(
      "web/GetFolderByServerRelativePath(decodedUrl='/Lib/My%20Folder')/Files/add(url='my%20file.pdf',overwrite=true)",
    );
  });
});
