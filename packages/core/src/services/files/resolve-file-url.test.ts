import { describe, expect, it, vi } from "vitest";
import type { SharePointRestClient } from "../../rest/client";
import type { LibraryContext } from "../../types/models";
import { resolveFileUrl } from "./resolve-file-url";

const library: LibraryContext = {
  listId: "11111111-1111-1111-1111-111111111111",
  listTitle: "Documents",
  rootFolderServerRelativeUrl: "/NPPDocumentDev",
  rootFolderUniqueId: "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
};

function restMock(response: unknown) {
  const get = vi.fn().mockResolvedValue(response);
  return { get, siteUrl: "https://contoso.sharepoint.com" } as unknown as SharePointRestClient;
}

describe("resolveFileUrl", () => {
  it("builds path GetFileById with ServerRelativeUrl select", async () => {
    const rest = restMock({ ServerRelativeUrl: "/NPPDocumentDev/Thư mục/a.txt" });
    const url = await resolveFileUrl(rest, async () => library, "aaef1547-b160-44dd-99b3-bfa7cd043dbc");
    expect(url).toBe("/NPPDocumentDev/Thư mục/a.txt");
    expect(rest.get).toHaveBeenCalledWith(
      "web/GetFileById('aaef1547-b160-44dd-99b3-bfa7cd043dbc')",
      { query: { $select: "ServerRelativeUrl" }, signal: undefined },
    );
  });

  it("rejects root alias — root has no file URL", async () => {
    const rest = restMock({});
    await expect(resolveFileUrl(rest, async () => library, "root")).rejects.toThrow(
      "Library root has no file URL",
    );
    expect(rest.get).not.toHaveBeenCalled();
  });

  it("rejects when ServerRelativeUrl missing", async () => {
    const rest = restMock({});
    await expect(
      resolveFileUrl(rest, async () => library, "aaef1547-b160-44dd-99b3-bfa7cd043dbc"),
    ).rejects.toThrow("not found");
  });
});
