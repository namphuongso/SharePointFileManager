import { describe, expect, it, vi } from "vitest";
import { mapGraphError, mapStatusToCode } from "./errors/map-graph-error";
import { SharePointErrorCode } from "./errors/sharepoint-error";
import { GraphClient } from "./graph/client";
import { mapDriveItem } from "./mappers/item";
import { mapGraphPerson, mapGraphUser, mapTypedEmail, toInviteRecipient } from "./mappers/person";
import { mapPermission } from "./mappers/permission";
import { extractMetadataFromListItem, mapListItemFields } from "./mappers/list-item";
import { resolveConfig } from "./config/resolve-config";
import { DriveService, findDriveByName } from "./services/drive";
import { SiteService } from "./services/site";
import { siteResourcePath } from "./graph/paths";
import { buildSearchKql } from "./services/search";
import { ActivityService } from "./services/activity";
import { ListItemService } from "./services/list-item";
import { itemsVisibleInFolder } from "./utils/list-visible-items";
import type { SharePointItem } from "./types/models";
import { canPerformItemAction, isOfficeOnlineFile } from "./utils/item-actions";
import {
  buildSharePointDocOpenUrl,
  isDirectFileDownloadUrl,
  resolveItemOpenUrl,
} from "./utils/sharepoint-open-url";

describe("mapStatusToCode", () => {
  it("maps HTTP statuses", () => {
    expect(mapStatusToCode(401)).toBe(SharePointErrorCode.Unauthorized);
    expect(mapStatusToCode(403)).toBe(SharePointErrorCode.Forbidden);
    expect(mapStatusToCode(404)).toBe(SharePointErrorCode.NotFound);
    expect(mapStatusToCode(409)).toBe(SharePointErrorCode.Conflict);
    expect(mapStatusToCode(429)).toBe(SharePointErrorCode.Throttled);
  });
});

describe("mapGraphError", () => {
  it("prefers Graph error codes", () => {
    const error = mapGraphError({
      status: 404,
      body: { error: { code: "itemNotFound", message: "Missing" } },
    });
    expect(error.code).toBe(SharePointErrorCode.NotFound);
    expect(error.message).toBe("Missing");
    expect(error.graphCode).toBe("itemNotFound");
  });

  it("parses Retry-After seconds", () => {
    const error = mapGraphError({ status: 429, retryAfter: "2" });
    expect(error.code).toBe(SharePointErrorCode.Throttled);
    expect(error.retryAfterMs).toBe(2000);
  });
});

describe("mapDriveItem", () => {
  it("maps a file", () => {
    const item = mapDriveItem({
      id: "1",
      name: "a.xlsx",
      size: 10,
      file: { mimeType: "application/vnd.ms-excel" },
      parentReference: { id: "parent", driveId: "drive" },
      "@microsoft.graph.downloadUrl": "https://example.com/file",
    });
    expect(item.type).toBe("file");
    expect(item.mimeType).toContain("excel");
    expect(item.downloadUrl).toBe("https://example.com/file");
  });

  it("maps a folder", () => {
    const item = mapDriveItem({
      id: "2",
      name: "Contracts",
      folder: { childCount: 3 },
    });
    expect(item.type).toBe("folder");
    expect(item.childCount).toBe(3);
  });

  it("maps sensitivity label from list item fields", () => {
    const item = mapDriveItem({
      id: "3",
      name: "secret.docx",
      file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
      listItem: {
        fields: {
          _ComplianceTag: "Confidential",
        },
      },
    });
    expect(item.sensitivityLabel).toBe("Confidential");
  });
});

describe("list item metadata mapper", () => {
  it("keeps custom internal names and lookup id helpers", () => {
    const mapped = extractMetadataFromListItem({
      id: "5",
      fields: {
        _CustomBusinessCode: "ABC-01",
        DepartmentLookupId: "12",
        CategoryStringId: "3;#Legal",
        _ComplianceTag: "Confidential",
        Modified: "2026-08-18T00:00:00Z",
      },
    });

    expect(mapped.metadata).toMatchObject({
      _CustomBusinessCode: "ABC-01",
      DepartmentLookupId: "12",
      CategoryStringId: "3;#Legal",
    });
    expect(mapped.metadata?._ComplianceTag).toBeUndefined();
    expect(mapped.metadata?.Modified).toBeUndefined();
  });

  it("normalizes multi-value and object field payloads without dropping labels", () => {
    const result = mapListItemFields("item-1", {
      fields: {
        Approvers: {
          results: [
            { LookupValue: "Ada Lovelace", Email: "ada@contoso.com" },
            { displayName: "Alan Turing" },
          ],
        },
        Tags: [{ Label: "Finance" }, { label: "Urgent" }],
      },
    });

    expect(result.fields.Approvers).toBe("Ada Lovelace, Alan Turing");
    expect(result.fields.Tags).toBe("Finance, Urgent");
  });
});

describe("mapPermission", () => {
  it("marks inherited permissions as not removable", () => {
    const permission = mapPermission({
      id: "p1",
      roles: ["read"],
      inheritedFrom: { id: "parent" },
      grantedToV2: { user: { displayName: "Ada", email: "ada@contoso.com" } },
    });
    expect(permission.inherited).toBe(true);
    expect(permission.canRemove).toBe(false);
    expect(permission.kind).toBe("user");
  });

  it("maps sharing links", () => {
    const permission = mapPermission({
      id: "p2",
      roles: ["write"],
      link: { type: "edit", scope: "organization", webUrl: "https://link" },
    });
    expect(permission.kind).toBe("link");
    expect(permission.canRemove).toBe(true);
    expect(permission.link?.scope).toBe("organization");
  });

  it("does not allow removing direct user/group ACL entries", () => {
    const userPermission = mapPermission({
      id: "p3",
      roles: ["write"],
      grantedToV2: { user: { id: "u1", displayName: "Ada" } },
    });
    const groupPermission = mapPermission({
      id: "p4",
      roles: ["read"],
      grantedToV2: { group: { id: "g1", displayName: "Legal Team" } },
    });
    expect(userPermission.kind).toBe("user");
    expect(groupPermission.kind).toBe("group");
    expect(userPermission.canRemove).toBe(false);
    expect(groupPermission.canRemove).toBe(false);
  });
});

describe("people mapper", () => {
  it("maps Graph people search results without using people-id as objectId", () => {
    const person = mapGraphPerson({
      id: "people-row-id",
      displayName: "Test 02",
      scoredEmailAddresses: [{ address: "test02@tcs.com.vn" }],
      personType: { class: "Person", subclass: "OrganizationUser" },
    });
    expect(person?.kind).toBe("user");
    expect(person?.email).toBe("test02@tcs.com.vn");
    expect(person?.objectId).toBeUndefined();
  });

  it("maps directory users with Azure AD object id", () => {
    const person = mapGraphUser({
      id: "aad-guid",
      displayName: "Test 02",
      mail: "test02@tcs.com.vn",
    });
    expect(person?.objectId).toBe("aad-guid");
    expect(toInviteRecipient(person!)?.objectId).toBe("aad-guid");
  });

  it("maps a typed email for external sharing", () => {
    const person = mapTypedEmail("omar_cou@hotmail.com");
    expect(person?.kind).toBe("email");
    expect(toInviteRecipient(person!)?.email).toBe("omar_cou@hotmail.com");
    expect(mapTypedEmail("not-an-email")).toBeUndefined();
  });
});

describe("resolveConfig", () => {
  it("fills defaults", () => {
    const config = resolveConfig({
      siteId: "site",
      tokenProvider: { getAccessToken: async () => "token" },
    });
    expect(config.rootItemId).toBe("root");
    expect(config.features.upload).toBe(true);
    expect(config.graphBaseUrl).toContain("graph.microsoft.com");
  });
});

describe("GraphClient", () => {
  it("builds versioned URLs with the configured API version", () => {
    const graph = new GraphClient({
      baseUrl: "https://graph.microsoft.com/v1.0",
      tokenProvider: { getAccessToken: async () => "token" },
      scopes: ["Sites.Read.All"],
    });
    expect(graph.apiUrl("/sites/site/drives")).toBe(
      "https://graph.microsoft.com/v1.0/sites/site/drives",
    );
  });

  it("normalizes a beta graphBaseUrl to v1.0", () => {
    const graph = new GraphClient({
      baseUrl: "https://graph.microsoft.com/beta",
      tokenProvider: { getAccessToken: async () => "token" },
      scopes: ["Sites.Read.All"],
    });
    expect(graph.apiUrl("/sites/site/drives")).toBe(
      "https://graph.microsoft.com/v1.0/sites/site/drives",
    );
  });

  it("retries once on 401 with forceRefresh", async () => {
    const getAccessToken = vi.fn(async ({ forceRefresh }: { forceRefresh?: boolean }) => {
      return forceRefresh ? "new-token" : "old-token";
    });
    const fetchImpl = vi.fn(async (_url: string, init?: RequestInit) => {
      const auth = new Headers(init?.headers).get("Authorization");
      if (auth === "Bearer old-token") {
        return new Response(JSON.stringify({ error: { code: "unauthenticated" } }), { status: 401 });
      }
      return new Response(JSON.stringify({ id: "ok" }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    });

    const client = new GraphClient({
      baseUrl: "https://graph.microsoft.com/v1.0",
      tokenProvider: { getAccessToken },
      scopes: ["Files.ReadWrite"],
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });

    const result = await client.get<{ id: string }>("/me");
    expect(result.id).toBe("ok");
    expect(getAccessToken).toHaveBeenCalledTimes(2);
    expect(getAccessToken.mock.calls[1]?.[0]).toMatchObject({ forceRefresh: true });
  });
});

describe("Graph-backed contracts", () => {
  it("uses documented managed properties for library search filters", () => {
    expect(buildSearchKql("budget", "https://contoso.sharepoint.com/sites/demo/Shared%20Documents", {
      modifiedAfter: "2025-01-01",
      modifiedBefore: "2025-12-31",
    })).toContain("LastModifiedTime>=2025-01-01");
  });

  it("maps v1 activity facets and timestamps", async () => {
    const fetchImpl = vi.fn(async (_url: string) => new Response(JSON.stringify({
      value: [{
        id: "a1",
        action: { edit: {} },
        activityDateTime: "2025-01-01T00:00:00Z",
        actor: { user: { id: "u1", displayName: "Ada" } },
        driveItem: { name: "plan.docx" },
      }],
    }), { status: 200, headers: { "Content-Type": "application/json" } }));
    const graph = new GraphClient({
      baseUrl: "https://graph.microsoft.com/v1.0",
      tokenProvider: { getAccessToken: async () => "token" },
      scopes: ["Sites.Read.All"],
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });
    const activities = await new ActivityService(graph, async () => "drive").list("item");
    expect(activities[0]).toMatchObject({ id: "a1", action: "edited", timestamp: "2025-01-01T00:00:00Z" });
    expect(String(fetchImpl.mock.calls[0]?.[0])).toContain("/drives/drive/items/item/activities");
  });

});

function jsonResponse(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

describe("ListItemService accessible library items", () => {
  it("reads driveItem from GET /sites/{site}/lists/{list}/items", async () => {
    const fetchImpl = vi.fn(async (url: string) => {
      const target = decodeURIComponent(String(url));
      expect(target).toContain("/sites/site-a/lists/list-a/items");
      expect(target).toContain("$expand=driveItem");
      return jsonResponse({
        value: [
          {
            id: "1",
            driveItem: {
              id: "folder-1",
              name: "Thư mục webmaster",
              folder: { childCount: 2 },
              parentReference: { driveId: "drive-a", id: "root-id" },
            },
          },
          {
            id: "2",
            driveItem: {
              id: "file-xlsx",
              name: "Workbook.xlsx",
              file: {},
              parentReference: { driveId: "drive-a", id: "folder-1" },
            },
          },
          {
            id: "3",
            driveItem: {
              id: "file-pptx",
              name: "webmaster.pptx",
              file: {},
              parentReference: { driveId: "drive-a", id: "folder-1" },
            },
          },
        ],
      });
    });

    const graph = new GraphClient({
      baseUrl: "https://graph.microsoft.com/v1.0",
      tokenProvider: { getAccessToken: async () => "token" },
      scopes: ["Sites.Read.All"],
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });

    const service = new ListItemService(graph, "site-a", async () => "drive-a", "list-a");
    const accessible = await service.listAccessibleDriveItems();
    const atRoot = itemsVisibleInFolder(accessible, "root");
    expect(atRoot).toHaveLength(1);
    expect(atRoot[0]?.name).toBe("Thư mục webmaster");

    const inFolder = itemsVisibleInFolder(accessible, "folder-1");
    expect(inFolder.map((item) => item.name).sort()).toEqual(["Workbook.xlsx", "webmaster.pptx"]);
  });
});

describe("itemsVisibleInFolder", () => {
  it("keeps files under their parent folder when the folder is visible", () => {
    const folder: SharePointItem = {
      id: "folder-1",
      name: "Thư mục webmaster",
      type: "folder",
      parentId: "root-id",
    };
    const xlsx: SharePointItem = {
      id: "file-xlsx",
      name: "Workbook.xlsx",
      type: "file",
      parentId: "folder-1",
    };
    const pptx: SharePointItem = {
      id: "file-pptx",
      name: "webmaster.pptx",
      type: "file",
      parentId: "folder-1",
    };
    expect(itemsVisibleInFolder([folder, xlsx, pptx], "root").map((item) => item.name)).toEqual([
      "Thư mục webmaster",
    ]);
    expect(
      itemsVisibleInFolder([folder, xlsx, pptx], "folder-1")
        .map((item) => item.name)
        .sort(),
    ).toEqual(["Workbook.xlsx", "webmaster.pptx"]);
  });
});

describe("sharepoint open url", () => {
  it("builds Doc.aspx url like SharePoint web", () => {
    const url = buildSharePointDocOpenUrl(
      "https://tcsvn.sharepoint.com/sites/eOfficeDev",
      "EF579525-42FC-41A3-A66F-F7B7B6599389",
      "Workbook.xlsx",
    );
    expect(url).toBe(
      "https://tcsvn.sharepoint.com/sites/eOfficeDev/_layouts/15/Doc.aspx?sourcedoc=%7BEF579525-42FC-41A3-A66F-F7B7B6599389%7D&file=Workbook.xlsx&action=default&mobileredirect=true",
    );
  });

  it("maps openUrl from sharepointIds for office files", () => {
    const item = mapDriveItem({
      id: "1",
      name: "Workbook.xlsx",
      file: { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
      webUrl: "https://tcsvn.sharepoint.com/sites/eOfficeDev/Shared Documents/Workbook.xlsx",
      sharepointIds: {
        siteUrl: "https://tcsvn.sharepoint.com/sites/eOfficeDev",
        listItemUniqueId: "EF579525-42FC-41A3-A66F-F7B7B6599389",
      },
    });
    expect(item.openUrl).toContain("/_layouts/15/Doc.aspx");
    expect(resolveItemOpenUrl(item)).toBe(item.openUrl);
    expect(isDirectFileDownloadUrl(item.webUrl!)).toBe(true);
  });
});

describe("item actions", () => {
  it("gates checkout actions from Graph checkout state", () => {
    const free: SharePointItem = {
      id: "1",
      name: "Workbook.xlsx",
      type: "file",
      capabilities: { isCheckedOut: false },
    };
    const locked: SharePointItem = {
      id: "2",
      name: "Workbook.xlsx",
      type: "file",
      capabilities: { isCheckedOut: true },
    };
    expect(canPerformItemAction(free, "download")).toBe(true);
    expect(canPerformItemAction(free, "checkout")).toBe(true);
    expect(canPerformItemAction(free, "checkin")).toBe(false);
    expect(canPerformItemAction(locked, "checkout")).toBe(false);
    expect(canPerformItemAction(locked, "checkin")).toBe(true);
    expect(canPerformItemAction(locked, "discardCheckout")).toBe(true);
  });

  it("detects office files by extension", () => {
    expect(
      isOfficeOnlineFile({ id: "1", name: "Report.xlsx", type: "file" }),
    ).toBe(true);
    expect(
      isOfficeOnlineFile({ id: "2", name: "Notes.txt", type: "file" }),
    ).toBe(false);
  });
});

describe("findDriveByName", () => {
  it("matches live Graph library names case-insensitively", () => {
    const drives = [
      { id: "1", name: "Documents" },
      { id: "2", name: "ISO Documents" },
    ];
    expect(findDriveByName(drives, "documents")?.id).toBe("1");
    expect(findDriveByName(drives, "ISO Documents")?.id).toBe("2");
    expect(findDriveByName(drives, "missing")).toBeUndefined();
  });
});

describe("DriveService.listDrives", () => {
  it("returns drives from GET /sites/{siteId}/drives", async () => {
    const fetchImpl = vi.fn(async (url: string) => {
      expect(String(url)).toContain("/sites/site-a/drives");
      return new Response(
        JSON.stringify({
          value: [
            { id: "d1", name: "Documents", driveType: "documentLibrary" },
            { name: "broken" },
          ],
        }),
        { status: 200, headers: { "Content-Type": "application/json" } },
      );
    });

    const graph = new GraphClient({
      baseUrl: "https://graph.microsoft.com/v1.0",
      tokenProvider: { getAccessToken: async () => "token" },
      scopes: ["Sites.ReadWrite.All"],
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });

    const drives = await new DriveService(graph, "site-a").listDrives();
    expect(drives).toEqual([
      { id: "d1", name: "Documents", webUrl: undefined, driveType: "documentLibrary" },
    ]);
  });

  it("uses hostname:path colon syntax for nested site resources", async () => {
    const fetchImpl = vi.fn(async (url: string) => {
      expect(String(url)).toBe(
        "https://graph.microsoft.com/v1.0/sites/tcsvn.sharepoint.com:/sites/eOffice:/drives",
      );
      return new Response(JSON.stringify({ value: [] }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    });

    const graph = new GraphClient({
      baseUrl: "https://graph.microsoft.com/v1.0",
      tokenProvider: { getAccessToken: async () => "token" },
      scopes: ["Sites.ReadWrite.All"],
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });

    await new DriveService(graph, "tcsvn.sharepoint.com:/sites/eOffice").listDrives();
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });
});

describe("siteIdentifierFromUrl", () => {
  it("builds Graph hostname:path from a SharePoint URL", async () => {
    const { siteIdentifierFromUrl } = await import("./services/site");
    expect(siteIdentifierFromUrl("https://tcsvn.sharepoint.com/sites/eOffice")).toBe(
      "tcsvn.sharepoint.com:/sites/eOffice",
    );
  });
});

describe("SiteService.listLists", () => {
  it("keeps the order returned by SharePoint instead of sorting by name", async () => {
    const fetchImpl = vi.fn(async (url: string) => {
      expect(String(url)).toContain("/sites/site-a/lists");
      return new Response(
        JSON.stringify({
          value: [
            { id: "1", name: "Tickets", displayName: "Tickets", list: { hidden: false } },
            { id: "2", name: "Devices", displayName: "Devices", list: { hidden: false } },
            { id: "3", name: "eNews", displayName: "eNews", list: { hidden: false } },
          ],
        }),
        { status: 200, headers: { "Content-Type": "application/json" } },
      );
    });

    const graph = new GraphClient({
      baseUrl: "https://graph.microsoft.com/v1.0",
      tokenProvider: { getAccessToken: async () => "token" },
      scopes: ["Sites.ReadWrite.All"],
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });

    const lists = await new SiteService(graph).listLists("site-a");
    expect(lists.map((list) => list.displayName)).toEqual(["Tickets", "Devices", "eNews"]);
  });
});

describe("isVisibleSharePointList", () => {
  it("hides system and catalog lists", async () => {
    const { isVisibleSharePointList, findListByName, mapGraphList } = await import("./services/site");
    expect(isVisibleSharePointList({ id: "1", displayName: "eNews", list: { hidden: false } })).toBe(true);
    expect(isVisibleSharePointList({ id: "2", displayName: "Style Library", list: { hidden: false } })).toBe(false);
    expect(isVisibleSharePointList({ id: "3", displayName: "eNews", list: { hidden: true } })).toBe(false);

    const lists = [
      mapGraphList({ id: "a", name: "eISODocuments", displayName: "eISODocuments", list: { template: "documentLibrary" } }),
      mapGraphList({ id: "b", name: "eNews", displayName: "eNews", list: { template: "genericList" } }),
    ].filter((item): item is NonNullable<typeof item> => Boolean(item));

    expect(findListByName(lists, "eNews")?.id).toBe("b");
  });
});

describe("siteResourcePath", () => {
  it("uses a slash for GUID site ids", () => {
    expect(siteResourcePath("c21db2f8-f48c-49b3-97e0-b9f83b4b73a0", "drives")).toBe(
      "/sites/c21db2f8-f48c-49b3-97e0-b9f83b4b73a0/drives",
    );
  });

  it("inserts a colon before nested resources for hostname:path ids", () => {
    expect(siteResourcePath("tcsvn.sharepoint.com:/sites/eOffice", "drives")).toBe(
      "/sites/tcsvn.sharepoint.com:/sites/eOffice:/drives",
    );
  });
});
