import { describe, expect, it, vi } from "vitest";
import type { SharePointItem } from "../../types/models";
import type { FolderService } from "../folder/folder";
import { collectFolderZipEntries } from "./collect-folder-zip-entries";

function item(partial: Pick<SharePointItem, "id" | "name" | "type">): SharePointItem {
  return {
    ...partial,
    lastModifiedDateTime: "2024-01-01T00:00:00Z",
  };
}

describe("collectFolderZipEntries", () => {
  it("walks nested folders via listChildren pages", async () => {
    const listChildren = vi.fn(async (folderId: string) => {
      if (folderId === "root-id") {
        return {
          items: [
            item({ id: "file-a", name: "a.txt", type: "file" }),
            item({ id: "sub-id", name: "Sub", type: "folder" }),
          ],
        };
      }
      if (folderId === "sub-id") {
        return {
          items: [item({ id: "file-b", name: "b.txt", type: "file" })],
        };
      }
      throw new Error(`unexpected folder ${folderId}`);
    });
    const folders = { listChildren } as unknown as FolderService;

    const result = await collectFolderZipEntries(folders, "root-id");
    expect(result.files).toEqual([
      { uniqueId: "file-a", relativePath: "a.txt" },
      { uniqueId: "file-b", relativePath: "Sub/b.txt" },
    ]);
    expect(result.emptyDirs).toEqual([]);
    expect(listChildren).toHaveBeenCalledWith("root-id", {
      top: 200,
      nextLink: undefined,
      signal: undefined,
    });
    expect(listChildren).toHaveBeenCalledWith("sub-id", {
      top: 200,
      nextLink: undefined,
      signal: undefined,
    });
  });

  it("follows nextLink until exhausted", async () => {
    const listChildren = vi
      .fn()
      .mockResolvedValueOnce({
        items: [item({ id: "file-a", name: "a.txt", type: "file" })],
        nextLink: "https://contoso.sharepoint.com/_api/next",
      })
      .mockResolvedValueOnce({
        items: [item({ id: "file-b", name: "b.txt", type: "file" })],
      });
    const folders = { listChildren } as unknown as FolderService;

    const result = await collectFolderZipEntries(folders, "root-id");
    expect(result.files).toEqual([
      { uniqueId: "file-a", relativePath: "a.txt" },
      { uniqueId: "file-b", relativePath: "b.txt" },
    ]);
    expect(listChildren).toHaveBeenNthCalledWith(2, "root-id", {
      top: 200,
      nextLink: "https://contoso.sharepoint.com/_api/next",
      signal: undefined,
    });
  });

  it("records empty nested directory", async () => {
    const listChildren = vi.fn(async (folderId: string) => {
      if (folderId === "root-id") {
        return { items: [item({ id: "empty-id", name: "Empty", type: "folder" })] };
      }
      return { items: [] };
    });
    const folders = { listChildren } as unknown as FolderService;

    const result = await collectFolderZipEntries(folders, "root-id");
    expect(result.files).toEqual([]);
    expect(result.emptyDirs).toEqual([{ relativePath: "Empty/" }]);
  });
});
