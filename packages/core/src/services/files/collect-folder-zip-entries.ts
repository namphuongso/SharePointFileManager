import type { SharePointItem } from "../../types/models";
import type { FolderService } from "../folder/folder";
import type { FolderZipDirEntry, FolderZipFileEntry } from "./folder-zip-types";

/** Trang lớn hơn browse UI — giảm số GET khi duyệt cây tải. */
const DOWNLOAD_LIST_PAGE_SIZE = 200;

export interface CollectFolderZipEntriesResult {
  files: FolderZipFileEntry[];
  emptyDirs: FolderZipDirEntry[];
}

/**
 * Duyệt cây folder bằng list items một cấp (`listChildren` + FileDirRef) — cùng pattern browse.
 * Forms đã loại trong mapRestListItem. Caller GET /$value từng file rồi zip.
 */
export async function collectFolderZipEntries(
  folders: FolderService,
  folderUniqueId: string,
  signal?: AbortSignal,
): Promise<CollectFolderZipEntriesResult> {
  const files: FolderZipFileEntry[] = [];
  const emptyDirs: FolderZipDirEntry[] = [];
  await walkFolder(folders, folderUniqueId, "", files, emptyDirs, signal);
  return { files, emptyDirs };
}

async function walkFolder(
  folders: FolderService,
  folderId: string,
  relativePrefix: string,
  files: FolderZipFileEntry[],
  emptyDirs: FolderZipDirEntry[],
  signal?: AbortSignal,
): Promise<void> {
  const children = await listAllChildren(folders, folderId, signal);

  if (children.length === 0) {
    if (relativePrefix) {
      emptyDirs.push({ relativePath: ensureTrailingSlash(relativePrefix) });
    }
    return;
  }

  for (const item of children) {
    const relativePath = joinZipPath(relativePrefix, item.name);
    if (item.type === "file") {
      files.push({ uniqueId: item.id, relativePath });
      continue;
    }
    await walkFolder(folders, item.id, relativePath, files, emptyDirs, signal);
  }
}

/** Lấy hết trang listChildren (nextLink) của một folder. */
async function listAllChildren(
  folders: FolderService,
  folderId: string,
  signal?: AbortSignal,
): Promise<SharePointItem[]> {
  const items: SharePointItem[] = [];
  let nextLink: string | undefined;
  do {
    const page = await folders.listChildren(folderId, {
      top: DOWNLOAD_LIST_PAGE_SIZE,
      nextLink,
      signal,
    });
    items.push(...page.items);
    nextLink = page.nextLink;
  } while (nextLink);
  return items;
}

function joinZipPath(prefix: string, name: string): string {
  return prefix ? `${prefix}/${name}` : name;
}

function ensureTrailingSlash(path: string): string {
  return path.endsWith("/") ? path : `${path}/`;
}
