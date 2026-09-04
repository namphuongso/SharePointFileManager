import type { SharePointRestClient } from "../../rest/client";
import { SharePointError, SharePointErrorCode } from "../../errors/sharepoint-error";
import type { LibraryContext } from "../../types/models";
import type { RestFolder } from "../../types/rest";
import { encodeServerRelativePathArg, normalizeGuid, requireUniqueId } from "../../utils";
import { relativeFolderSegments } from "./relative-folder-segments";

/** Một cấp breadcrumb (không gồm root UI) — UniqueId + Name. */
export interface FolderBreadcrumbSegment {
  id: string;
  name: string;
}

/**
 * UniqueId folder → chuỗi breadcrumb dưới root (GetFolderById + ancestors theo path).
 * Root / "root" → []. Dùng khi hydrate URL ?folder= sau F5.
 */
export async function resolveFolderBreadcrumb(
  rest: SharePointRestClient,
  getLibrary: () => Promise<LibraryContext>,
  folderId: string,
  signal?: AbortSignal,
): Promise<FolderBreadcrumbSegment[]> {
  const library = await getLibrary();
  if (folderId === "root" || normalizeGuid(folderId) === normalizeGuid(library.rootFolderUniqueId)) {
    return [];
  }

  const folder = await rest.get<RestFolder>(
    `web/GetFolderById('${normalizeGuid(folderId)}')`,
    {
      query: { $select: "UniqueId,Name,ServerRelativeUrl" },
      signal,
    },
  );

  if (!folder.ServerRelativeUrl || !folder.UniqueId || !folder.Name) {
    throw new SharePointError({
      code: SharePointErrorCode.NotFound,
      message: `Folder ${folderId} was not found`,
    });
  }

  const parts = relativeFolderSegments(
    library.rootFolderServerRelativeUrl,
    folder.ServerRelativeUrl,
  );
  if (parts === null) {
    throw new SharePointError({
      code: SharePointErrorCode.NotFound,
      message: `Folder ${folderId} is outside the current library`,
    });
  }
  if (parts.length === 0) return [];

  const rootUrl = library.rootFolderServerRelativeUrl.replace(/\/$/, "");
  const segments: FolderBreadcrumbSegment[] = [];

  let cumulative = rootUrl;
  for (let index = 0; index < parts.length; index++) {
    cumulative = `${cumulative}/${parts[index]}`;
    const isLeaf = index === parts.length - 1;
    if (isLeaf) {
      segments.push({
        id: requireUniqueId(folder.UniqueId, "folder"),
        name: folder.Name,
      });
      break;
    }

    const ancestor = await rest.get<RestFolder>(
      `web/GetFolderByServerRelativePath(decodedUrl='${encodeServerRelativePathArg(cumulative)}')`,
      {
        query: { $select: "UniqueId,Name" },
        signal,
      },
    );
    if (!ancestor.UniqueId || !ancestor.Name) {
      throw new SharePointError({
        code: SharePointErrorCode.NotFound,
        message: `Ancestor folder not found at ${cumulative}`,
      });
    }
    segments.push({
      id: requireUniqueId(ancestor.UniqueId, "folder"),
      name: ancestor.Name,
    });
  }

  return segments;
}
