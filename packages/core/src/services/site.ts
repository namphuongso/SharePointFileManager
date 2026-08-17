import type { GraphClient } from "../graph/client";
import { siteResourcePath } from "../graph/paths";
import type { GraphCollection } from "../mappers/item";
import type { SharePointListInfo, SiteInfo } from "../types/models";

interface GraphSite {
  id?: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
}

interface GraphList {
  id?: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
  system?: boolean;
  list?: {
    hidden?: boolean;
    template?: string;
  };
  drive?: {
    id?: string;
  };
}

const HIDDEN_LIST_NAMES = new Set(
  [
    "form templates",
    "style library",
    "master page gallery",
    "solution gallery",
    "theme gallery",
    "web part gallery",
    "list template gallery",
    "user information list",
    "taxonomyhiddenlist",
    "converted forms",
    "device channels",
    "composed looks",
    "maintenance log library",
    "sharepointhomecachelist",
    "access requests",
    "cache profiles",
    "content type sync log",
    "microfeed",
    "wfpub",
    "appdata",
    "workflow history",
    "reusable content",
  ].map((name) => name.toLowerCase()),
);

export function siteIdentifierFromUrl(siteUrl: string): string {
  const parsed = new URL(siteUrl.trim());
  const path = parsed.pathname.replace(/\/$/, "") || "/";
  return `${parsed.hostname}:${path}`;
}

export function isVisibleSharePointList(list: GraphList): boolean {
  if (!list.id || list.system || list.list?.hidden) return false;
  const name = (list.displayName || list.name || "").trim().toLowerCase();
  if (!name) return false;
  if (HIDDEN_LIST_NAMES.has(name)) return false;
  return true;
}

export function mapGraphList(list: GraphList): SharePointListInfo | undefined {
  if (!isVisibleSharePointList(list) || !list.id) return undefined;
  const name = (list.name || list.displayName || "").trim();
  const displayName = (list.displayName || list.name || "").trim();
  if (!name && !displayName) return undefined;
  const template = list.list?.template;
  return {
    id: list.id,
    name: name || displayName,
    displayName: displayName || name,
    webUrl: list.webUrl,
    template,
    hasDrive: Boolean(list.drive?.id) || template === "documentLibrary" || template === "webPageLibrary",
  };
}

export function findListByName(lists: SharePointListInfo[], libraryName: string): SharePointListInfo | undefined {
  const needle = libraryName.trim().toLowerCase();
  if (!needle) return undefined;
  return lists.find(
    (list) => list.displayName.trim().toLowerCase() === needle || list.name.trim().toLowerCase() === needle,
  );
}

export class SiteService {
  constructor(private readonly graph: GraphClient) {}

  async getByUrl(siteUrl: string, signal?: AbortSignal): Promise<SiteInfo> {
    const identifier = siteIdentifierFromUrl(siteUrl);
    const site = await this.graph.get<GraphSite>(siteResourcePath(identifier), { signal });
    if (!site.id) {
      throw new Error("SharePoint site was not returned by Microsoft Graph");
    }
    return {
      id: site.id,
      name: site.name,
      displayName: site.displayName,
      webUrl: site.webUrl,
    };
  }

  async listLists(siteId: string, signal?: AbortSignal): Promise<SharePointListInfo[]> {
    const lists: SharePointListInfo[] = [];
    let path: string | undefined = siteResourcePath(siteId, "lists");
    let absoluteUrl = false;

    while (path) {
      const result: GraphCollection<GraphList> = await this.graph.get<GraphCollection<GraphList>>(path, {
        signal,
        absoluteUrl,
        query: absoluteUrl ? undefined : { $top: 200, $expand: "drive" },
      });

      for (const raw of result.value ?? []) {
        const mapped = mapGraphList(raw);
        if (mapped) lists.push(mapped);
      }

      path = result["@odata.nextLink"];
      absoluteUrl = true;
    }

    return lists;
  }
}
