import { describe, expect, it, vi } from "vitest";
import { SharePointError, SharePointErrorCode } from "../../errors/sharepoint-error";
import type { LibraryContext } from "../../types/models";
import { DeleteItemService } from "./delete-item-service";

const library: LibraryContext = {
  listId: "11111111-1111-1111-1111-111111111111",
  listTitle: "Docs",
  rootFolderServerRelativeUrl: "/sites/demo/Docs",
  rootFolderUniqueId: "22222222-2222-2222-2222-222222222222",
};

describe("DeleteItemService", () => {
  it("POSTs recycle path for a file UniqueId", async () => {
    const post = vi.fn().mockResolvedValue(undefined);
    const service = new DeleteItemService(
      { post } as never,
      async () => library,
    );

    await service.delete("file", "aaef1547-b160-44dd-99b3-bfa7cd043dbc");

    expect(post).toHaveBeenCalledWith(
      "web/lists(guid'11111111-1111-1111-1111-111111111111')/GetItemByUniqueId(guid'aaef1547-b160-44dd-99b3-bfa7cd043dbc')/recycle()",
      expect.objectContaining({ signal: undefined }),
    );
  });

  it("rejects library root", async () => {
    const post = vi.fn();
    const service = new DeleteItemService(
      { post } as never,
      async () => library,
    );

    await expect(service.delete("folder", "root")).rejects.toMatchObject({
      code: SharePointErrorCode.Unsupported,
    });
    await expect(
      service.delete("folder", library.rootFolderUniqueId),
    ).rejects.toBeInstanceOf(SharePointError);
    expect(post).not.toHaveBeenCalled();
  });
});
