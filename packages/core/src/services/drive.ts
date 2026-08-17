import type { GraphClient } from "../graph/client";
import { siteResourcePath } from "../graph/paths";
import type { GraphCollection } from "../mappers/item";
import type { DriveInfo } from "../types/models";

interface GraphDrive {
  id?: string;
  name?: string;
  webUrl?: string;
  driveType?: string;
}

export function mapGraphDrive(drive: GraphDrive): DriveInfo | undefined {
  if (!drive.id) return undefined;
  return {
    id: drive.id,
    name: drive.name,
    webUrl: drive.webUrl,
    driveType: drive.driveType,
  };
}

export function findDriveByName(drives: DriveInfo[], libraryName: string): DriveInfo | undefined {
  const needle = libraryName.trim().toLowerCase();
  if (!needle) return undefined;
  return drives.find((drive) => drive.name?.trim().toLowerCase() === needle);
}

export class DriveService {
  constructor(
    private readonly graph: GraphClient,
    private readonly siteId: string,
  ) {}

  async listDrives(signal?: AbortSignal): Promise<DriveInfo[]> {
    const drives: DriveInfo[] = [];
    let path: string | undefined = siteResourcePath(this.siteId, "drives");
    let absoluteUrl = false;

    while (path) {
      const result: GraphCollection<GraphDrive> = await this.graph.get<GraphCollection<GraphDrive>>(path, {
        signal,
        absoluteUrl,
      });
      for (const drive of result.value ?? []) {
        const mapped = mapGraphDrive(drive);
        if (mapped) drives.push(mapped);
      }
      path = result["@odata.nextLink"];
      absoluteUrl = true;
    }

    return drives;
  }

  async getDefaultDrive(signal?: AbortSignal): Promise<DriveInfo> {
    const drive = await this.graph.get<GraphDrive>(siteResourcePath(this.siteId, "drive"), { signal });
    if (!drive.id) {
      throw new Error("Default drive for site was not returned by Microsoft Graph");
    }
    return {
      id: drive.id,
      name: drive.name,
      webUrl: drive.webUrl,
      driveType: drive.driveType,
    };
  }

  async getDrive(driveId: string, signal?: AbortSignal): Promise<DriveInfo> {
    const drive = await this.graph.get<GraphDrive>(`/drives/${driveId}`, { signal });
    if (!drive.id) {
      throw new Error("Drive was not returned by Microsoft Graph");
    }
    return {
      id: drive.id,
      name: drive.name,
      webUrl: drive.webUrl,
      driveType: drive.driveType,
    };
  }

  async getDriveByList(listId: string, signal?: AbortSignal): Promise<DriveInfo> {
    const drive = await this.graph.get<GraphDrive>(siteResourcePath(this.siteId, `lists/${listId}/drive`), {
      signal,
    });
    if (!drive.id) {
      throw new Error("Drive for SharePoint list was not returned by Microsoft Graph");
    }
    return {
      id: drive.id,
      name: drive.name,
      webUrl: drive.webUrl,
      driveType: drive.driveType,
    };
  }
}
