import type { GraphClient } from "../graph/client";
import { pathByNameUrl } from "../graph/paths";
import { mapDriveItem, type GraphDriveItem } from "../mappers/item";
import type { SharePointItem, UploadOptions } from "../types/models";

const SMALL_FILE_LIMIT = 4 * 1024 * 1024;
const CHUNK_SIZE = 320 * 1024 * 10; // 3.2 MiB, multiple of 320 KiB

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
    const session = await this.graph.post<{ uploadUrl: string }>(
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
    let offset = 0;
    let lastItem: GraphDriveItem | undefined;

    try {
      while (offset < total) {
        if (options.signal?.aborted) {
          await this.cancelSession(session.uploadUrl);
          throw options.signal.reason;
        }

        const end = Math.min(offset + CHUNK_SIZE, total);
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
        }
      }
    } catch (error) {
      await this.cancelSession(session.uploadUrl);
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
}

async function toUint8Array(content: Blob | ArrayBuffer | Uint8Array): Promise<Uint8Array> {
  if (content instanceof Uint8Array) return content;
  if (content instanceof ArrayBuffer) return new Uint8Array(content);
  const buffer = await content.arrayBuffer();
  return new Uint8Array(buffer);
}
