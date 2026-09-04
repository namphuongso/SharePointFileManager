import { zipSync } from "fflate";
import type { SharePointRestClient } from "../../rest/client";
import { SharePointError, SharePointErrorCode } from "../../errors/sharepoint-error";
import type { FolderService } from "../folder/folder";
import { fileDownloadPath } from "./file-download-path";
import { collectFolderZipEntries } from "./collect-folder-zip-entries";

/** Tiến trình tải folder — UI giữ toast sticky + cập nhật subtitle. */
export type FolderZipPhase = "listing" | "downloading" | "zipping";

export interface FolderZipProgress {
  phase: FolderZipPhase;
  /** Số file đã GET /$value (phase downloading). */
  done: number;
  /** Tổng file cần tải (0 khi đang listing). */
  total: number;
}

export interface BuildFolderZipOptions {
  signal?: AbortSignal;
  /** Tên folder từ UI (SharePointItem.name) — đặt tên .zip và thư mục gốc trong archive. */
  folderName?: string;
  onProgress?: (progress: FolderZipProgress) => void;
}

export interface BuildFolderZipResult {
  blob: Blob;
  /** Tên file đề xuất: `{folderName}.zip`. */
  fileName: string;
}

/**
 * Tải cả cây folder: listChildren đệ quy → GET /$value từng file → zip (fflate).
 * Không có endpoint SharePoint tải folder một phát; zip phía client.
 */
export async function buildFolderZip(
  rest: SharePointRestClient,
  folders: FolderService,
  folderUniqueId: string,
  options?: BuildFolderZipOptions,
): Promise<BuildFolderZipResult> {
  options?.onProgress?.({ phase: "listing", done: 0, total: 0 });

  const { files, emptyDirs } = await collectFolderZipEntries(
    folders,
    folderUniqueId,
    options?.signal,
  );

  const rootPrefix = sanitizeZipBaseName(options?.folderName?.trim() || "folder");
  const entries: Record<string, Uint8Array> = {};
  for (const dir of emptyDirs) {
    entries[`${rootPrefix}/${dir.relativePath}`] = new Uint8Array(0);
  }

  const total = files.length;
  options?.onProgress?.({ phase: "downloading", done: 0, total });

  for (let i = 0; i < files.length; i++) {
    const file = files[i]!;
    const blob = await rest.getBlob(fileDownloadPath(file.uniqueId), {
      signal: options?.signal,
    });
    entries[`${rootPrefix}/${file.relativePath}`] = new Uint8Array(await blob.arrayBuffer());
    options?.onProgress?.({ phase: "downloading", done: i + 1, total });
  }

  if (Object.keys(entries).length === 0) {
    entries[`${rootPrefix}/`] = new Uint8Array(0);
  }

  options?.onProgress?.({ phase: "zipping", done: total, total });

  let zipped: Uint8Array;
  try {
    zipped = zipSync(entries, { level: 6 });
  } catch (cause) {
    throw new SharePointError({
      code: SharePointErrorCode.Unknown,
      message: "Failed to build folder zip",
      cause,
    });
  }

  const copy = new Uint8Array(zipped.byteLength);
  copy.set(zipped);
  return {
    blob: new Blob([copy.buffer], { type: "application/zip" }),
    fileName: `${rootPrefix}.zip`,
  };
}

/** Bỏ ký tự không hợp lệ trên tên file zip (Windows / macOS). */
function sanitizeZipBaseName(name: string): string {
  const cleaned = name.replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_").trim();
  return cleaned || "folder";
}
