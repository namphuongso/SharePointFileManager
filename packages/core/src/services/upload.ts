import type { GraphClient } from "../graph/client";
import { childrenUrl } from "../graph/paths";
import { pathByNameUrl } from "../graph/paths";
import { mapDriveItem, type GraphDriveItem } from "../mappers/item";
import type { SharePointItem, UploadOptions } from "../types/models";

const SMALL_FILE_LIMIT = 4 * 1024 * 1024;
const DEFAULT_CHUNK_SIZE = 320 * 1024 * 32; // 10 MiB, Microsoft-recommended stable chunk size

export class UploadService {
  constructor(
    private readonly graph: GraphClient,
    private readonly getDriveId: () => Promise<string>,
  ) {}

  async upload(options: UploadOptions): Promise<SharePointItem> {
    const bytes = await toUint8Array(options.content);
    if (bytes.byteLength <= SMALL_FILE_LIMIT) {
      return this.uploadSmallFile(options, bytes);
    }
    return this.uploadLargeFile(options, bytes);
  }

  async createPlaceholder(parentId: string, fileName: string, signal?: AbortSignal): Promise<SharePointItem> {
    const driveId = await this.getDriveId();
    const item = await this.graph.post<GraphDriveItem>(
      childrenUrl(driveId, parentId),
      {
        name: fileName,
        file: {},
        "@microsoft.graph.conflictBehavior": "rename",
      },
      { signal },
    );
    return mapDriveItem(item, driveId);
  }

  private async uploadSmallFile(options: UploadOptions, bytes: Uint8Array): Promise<SharePointItem> {
    const driveId = await this.getDriveId();
    const path = pathByNameUrl(driveId, options.parentId, options.fileName, "content");
    const item = await this.graph.put<GraphDriveItem>(path, bytes, {
      headers: {
        "Content-Type": "application/octet-stream",
      },
      query: {
        "@microsoft.graph.conflictBehavior": options.conflictBehavior ?? "rename",
      },
      signal: options.signal,
    });
    options.onProgress?.({
      bytesUploaded: bytes.byteLength,
      totalBytes: bytes.byteLength,
      percent: 100,
    });
    return mapDriveItem(item, driveId);
  }

  private async uploadLargeFile(options: UploadOptions, bytes: Uint8Array): Promise<SharePointItem> {
    const driveId = await this.getDriveId();
    const session = options.resume?.uploadUrl
      ? { uploadUrl: options.resume.uploadUrl }
      : await this.graph.post<{ uploadUrl: string }>(
      pathByNameUrl(driveId, options.parentId, options.fileName, "createUploadSession"),
      {
        item: {
          "@microsoft.graph.conflictBehavior": options.conflictBehavior ?? "rename",
          name: options.fileName,
        },
      },
      { signal: options.signal },
    );

    const total = bytes.byteLength;
    const chunkSize = options.chunkSize ?? DEFAULT_CHUNK_SIZE;
    if (chunkSize <= 0 || chunkSize >= 60 * 1024 * 1024 || chunkSize % (320 * 1024) !== 0) {
      throw new Error("chunkSize must be a positive multiple of 320 KiB and smaller than 60 MiB");
    }
    let offset = options.resume?.nextOffset;
    if (offset === undefined && options.resume?.uploadUrl) {
      offset = await this.getNextOffset(session.uploadUrl, options.signal);
    }
    offset ??= 0;
    if (offset < 0 || offset > total) throw new Error("resume nextOffset is outside the file range");
    let lastItem: GraphDriveItem | undefined;

    try {
      while (offset < total) {
        if (options.signal?.aborted) {
          await this.cancelSession(session.uploadUrl);
          throw new DOMException("Upload was cancelled", "AbortError");
        }

        const end = Math.min(offset + chunkSize, total);
        const chunk = bytes.slice(offset, end);
        const response = await this.graph.request<Response>({
          path: session.uploadUrl,
          absoluteUrl: true,
          skipAuth: true,
          method: "PUT",
          body: chunk,
          headers: {
            "Content-Length": String(chunk.byteLength),
            "Content-Range": `bytes ${offset}-${end - 1}/${total}`,
          },
          raw: true,
          signal: options.signal,
        });

        offset = end;
        options.onProgress?.({
          bytesUploaded: offset,
          totalBytes: total,
          percent: Math.round((offset / total) * 100),
        });

        if (response.status === 201 || response.status === 200) {
          lastItem = (await response.json()) as GraphDriveItem;
        } else if (response.status === 202) {
          const ranges = (await response.json()) as { nextExpectedRanges?: string[] };
          const next = ranges.nextExpectedRanges?.[0]?.split("-")[0];
          if (next && Number.isFinite(Number(next))) offset = Number(next);
        }
      }
    } catch (error) {
      // Keep a valid upload session so callers can query status and resume after transient failures.
      throw error;
    }

    if (!lastItem?.id) {
      throw new Error("Upload session completed without a driveItem response");
    }
    return mapDriveItem(lastItem, driveId);
  }

  private async cancelSession(uploadUrl: string): Promise<void> {
    try {
      await this.graph.request({
        path: uploadUrl,
        absoluteUrl: true,
        skipAuth: true,
        method: "DELETE",
        raw: true,
      });
    } catch {
      // Best-effort cancel.
    }
  }

  private async getNextOffset(uploadUrl: string, signal?: AbortSignal): Promise<number> {
    const status = await this.graph.request<{ nextExpectedRanges?: string[] }>({
      path: uploadUrl,
      absoluteUrl: true,
      skipAuth: true,
      method: "GET",
      signal,
    });
    const firstRange = status.nextExpectedRanges?.[0];
    const firstOffset = firstRange?.split("-")[0];
    if (!firstOffset || !/^\d+$/.test(firstOffset)) {
      throw new Error("Upload session did not return a valid nextExpectedRanges value");
    }
    return Number(firstOffset);
  }
}

async function toUint8Array(content: Blob | ArrayBuffer | Uint8Array): Promise<Uint8Array> {
  if (content instanceof Uint8Array) return content;
  if (content instanceof ArrayBuffer) return new Uint8Array(content);
  const buffer = await content.arrayBuffer();
  return new Uint8Array(buffer);
}
