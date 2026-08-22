import type { SharePointItem } from "../types/models";

export interface RestFile {
  UniqueId?: string;
  Name?: string;
  Length?: number;
  TimeLastModified?: string;
}

export interface RestFolder {
  UniqueId?: string;
  Name?: string;
  ServerRelativeUrl?: string;
  TimeLastModified?: string;
}

export function mapRestFile(file: RestFile): SharePointItem {
  const name = file.Name ?? "file";
  const id = file.UniqueId;
  if (!id) throw new Error("SharePoint file is missing UniqueId");
  return {
    id,
    name,
    type: "file",
    size: file.Length,
    lastModifiedDateTime: file.TimeLastModified,
  };
}

export function mapRestFolder(folder: RestFolder): SharePointItem {
  const name = folder.Name ?? "folder";
  const id = folder.UniqueId;
  if (!id) throw new Error("SharePoint folder is missing UniqueId");
  return {
    id,
    name,
    type: "folder",
    lastModifiedDateTime: folder.TimeLastModified,
  };
}
