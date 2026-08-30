import type { SharePointRestClient } from "../../rest/client";
import { SharePointError, SharePointErrorCode } from "../../errors/sharepoint-error";
import type { LibraryContext } from "../../types/models";
import type { RestFileUrl } from "../../types/rest";

/** Sự kiện click để mở: file mới cần URL; folder là điều hướng UI không REST. */
export type OpenUrlKind = "file" | "folder";

/**
 * ServerRelativeUrl của file theo UniqueId — phục vụ mở ?web=1.
 * Folder không đi qua đây (điều hướng breadcrumb trong UI).
 * @see https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-rest-reference/jj245549(v=office.15)
 */
export async function resolveFileUrl(
  rest: SharePointRestClient,
  getLibrary: () => Promise<LibraryContext>,
  fileId: string,
  signal?: AbortSignal,
): Promise<string> {
  const library = await getLibrary();
  if (fileId === "root" || fileId === library.rootFolderUniqueId) {
    throw new SharePointError({
      code: SharePointErrorCode.NotFound,
      message: "Library root has no file URL",
    });
  }

  const file = await rest.get<RestFileUrl>(`web/GetFileById('${fileId}')`, {
    query: { $select: "ServerRelativeUrl" },
    signal,
  });
  if (!file.ServerRelativeUrl) {
    throw new SharePointError({
      code: SharePointErrorCode.NotFound,
      message: `File ${fileId} was not found`,
    });
  }
  return file.ServerRelativeUrl;
}
