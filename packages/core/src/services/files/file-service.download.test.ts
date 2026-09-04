import { describe, expect, it, vi } from "vitest";
import type { SharePointRestClient } from "../../rest/client";
import type { LibraryContext } from "../../types/models";
import type { FolderService } from "../folder/folder";
import { FileService } from "./file-service";

const library: LibraryContext = {
  listId: "11111111-1111-1111-1111-111111111111",
  listTitle: "Documents",
  rootFolderServerRelativeUrl: "/NPPDocumentDev",
  rootFolderUniqueId: "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
};

describe("FileService.download", () => {
  it("GETs GetFileById /$value as blob", async () => {
    const blob = new Blob(["hello"], { type: "text/plain" });
    const getBlob = vi.fn().mockResolvedValue(blob);
    const rest = {
      getBlob,
      siteUrl: "https://contoso.sharepoint.com",
    } as unknown as SharePointRestClient;
    const folders = { listChildren: vi.fn() } as unknown as FolderService;
    const files = new FileService(rest, async () => library, folders);

    const result = await files.download("aaef1547-b160-44dd-99b3-bfa7cd043dbc");
    expect(result.blob).toBe(blob);
    expect(getBlob).toHaveBeenCalledWith(
      "web/GetFileById('aaef1547-b160-44dd-99b3-bfa7cd043dbc')/$value",
      { signal: undefined },
    );
  });

  it("rejects library root — no file content", async () => {
    const getBlob = vi.fn();
    const rest = { getBlob } as unknown as SharePointRestClient;
    const folders = { listChildren: vi.fn() } as unknown as FolderService;
    const files = new FileService(rest, async () => library, folders);

    await expect(files.download("root")).rejects.toThrow("Library root has no file content");
    await expect(files.download(library.rootFolderUniqueId)).rejects.toThrow(
      "Library root has no file content",
    );
    expect(getBlob).not.toHaveBeenCalled();
  });
});
