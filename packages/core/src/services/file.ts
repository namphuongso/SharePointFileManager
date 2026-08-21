import type { GraphClient } from "../graph/client";
import { itemUrl } from "../graph/paths";
import type { FolderService } from "./folder";
import type { FileVersion, PreviewInfo, SharePointItem } from "../types/models";

export class FileService {
  constructor(
    private readonly graph: GraphClient,
    private readonly getDriveId: () => Promise<string>,
    private readonly folders: FolderService,
  ) {}

  get(itemId: string, signal?: AbortSignal): Promise<SharePointItem> {
    return this.folders.get(itemId, { signal, expandListItem: true });
  }

  async download(
    itemId: string,
    signal?: AbortSignal,
  ): Promise<{ blob: Blob; fileName: string; mimeType?: string }> {
    return this.downloadContent(itemId, "content", signal);
  }

  async downloadVersion(
    itemId: string,
    versionId: string,
    signal?: AbortSignal,
  ): Promise<{ blob: Blob; fileName: string; mimeType?: string }> {
    return this.downloadContent(itemId, `versions/${versionId}/content`, signal);
  }

  private async downloadContent(
    itemId: string,
    contentPath: string,
    signal?: AbortSignal,
  ): Promise<{ blob: Blob; fileName: string; mimeType?: string }> {
    const item = await this.get(itemId, signal);
    const driveId = await this.getDriveId();

    let blob: Blob;
    if (contentPath === "content" && item.downloadUrl) {
      const response = await this.graph.request<Response>({
        path: item.downloadUrl,
        absoluteUrl: true,
        skipAuth: true,
        raw: true,
        signal,
      });
      blob = await response.blob();
    } else {
      const response = await this.graph.request<Response>({
        path: `${itemUrl(driveId, itemId)}/${contentPath}`,
        raw: true,
        signal,
      });
      blob = await response.blob();
    }

    return { blob, fileName: item.name, mimeType: item.mimeType };
  }

  async getVersions(itemId: string, signal?: AbortSignal): Promise<FileVersion[]> {
    const driveId = await this.getDriveId();
    const result = await this.graph.get<{
      value?: Array<{
        id?: string;
        lastModifiedDateTime?: string;
        size?: number;
        lastModifiedBy?: { user?: { id?: string; displayName?: string; email?: string } };
      }>;
    }>(`${itemUrl(driveId, itemId)}/versions`, { signal });

    return (result.value ?? []).map((version) => ({
      id: version.id ?? "",
      lastModifiedDateTime: version.lastModifiedDateTime,
      size: version.size,
      lastModifiedBy: version.lastModifiedBy?.user
        ? {
            id: version.lastModifiedBy.user.id,
            displayName: version.lastModifiedBy.user.displayName,
            email: version.lastModifiedBy.user.email,
          }
        : undefined,
    }));
  }

  async restoreVersion(itemId: string, versionId: string, signal?: AbortSignal): Promise<void> {
    const driveId = await this.getDriveId();
    await this.graph.post(`${itemUrl(driveId, itemId)}/versions/${versionId}/restoreVersion`, undefined, {
      signal,
    });
  }

  async preview(
    itemId: string,
    options?: { driveId?: string; signal?: AbortSignal },
  ): Promise<PreviewInfo> {
    const driveId = options?.driveId ?? (await this.getDriveId());
    return this.graph.post<PreviewInfo>(`${itemUrl(driveId, itemId)}/preview`, {}, { signal: options?.signal });
  }
}
