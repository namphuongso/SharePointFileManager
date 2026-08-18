import { describe, expect, it, vi } from "vitest";
import { mapGraphError, mapStatusToCode } from "./errors/map-graph-error";
import { SharePointErrorCode } from "./errors/sharepoint-error";
import { GraphClient } from "./graph/client";
import { mapDriveItem } from "./mappers/item";
import { mapGraphPerson, mapGraphUser, mapTypedEmail, toInviteRecipient } from "./mappers/person";
import { mapPermission } from "./mappers/permission";
import { resolveConfig } from "./config/resolve-config";
import { DriveService, findDriveByName } from "./services/drive";
import { SiteService } from "./services/site";
import { siteResourcePath } from "./graph/paths";
import { buildSearchKql } from "./services/search";
import { ActivityService } from "./services/activity";
import { DeltaService } from "./services/delta";

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
      scopes: ["Files.Read"],
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
      scopes: ["Files.Read"],
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });
    const activities = await new ActivityService(graph, async () => "drive").list("item");
    expect(activities[0]).toMatchObject({ id: "a1", action: "edit", timestamp: "2025-01-01T00:00:00Z" });
    expect(String(fetchImpl.mock.calls[0]?.[0])).toContain("/drives/drive/items/item/activities");
  });

  it("restarts a delta sync when Graph invalidates a saved token", async () => {
    let calls = 0;
    const fetchImpl = vi.fn(async (url: string) => {
      calls += 1;
      if (calls === 1) return new Response(JSON.stringify({ error: { code: "resyncRequired" } }), { status: 410 });
      expect(url).toContain("/drives/drive/root/delta");
      return new Response(JSON.stringify({ value: [], "@odata.deltaLink": "https://delta/new" }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    });
    const graph = new GraphClient({
      baseUrl: "https://graph.microsoft.com/v1.0",
      tokenProvider: { getAccessToken: async () => "token" },
      scopes: ["Files.Read"],
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });
    const result = await new DeltaService(graph).sync("drive", "root", { deltaLink: "https://delta/expired" });
    expect(result.deltaLink).toBe("https://delta/new");
    expect(calls).toBe(2);
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
