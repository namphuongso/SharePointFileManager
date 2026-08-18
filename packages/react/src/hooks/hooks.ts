import { useInfiniteQuery, useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import {
  SharePointClient,
  decodeLibraryPage,
  findDriveByName,
  findListByName,
  isSharePointError,
  SharePointErrorCode,
  siteIdentifierFromUrl,
  type CreateLinkOptions,
  type DriveInfo,
  type InviteOptions,
  type SearchFilters,
  type SearchScope,
  type SharePointConfig,
  type SharePointItem,
  type SharePointListInfo,
  type SiteInfo,
  type UploadProgress,
} from "@namphuongso/sharepoint-file-manager-core";
import { useSharePoint } from "../provider/context";
import { queryKeys } from "./queryKeys";

export function useSharePointSite(
  config:
    | (Pick<SharePointConfig, "tokenProvider" | "scopes" | "graphBaseUrl"> & { siteUrl: string })
    | undefined,
) {
  return useQuery({
    queryKey: queryKeys.site(config?.siteUrl ?? ""),
    enabled: Boolean(config?.siteUrl && config?.tokenProvider),
    queryFn: async ({ signal }): Promise<SiteInfo> => {
      const client = new SharePointClient({
        siteId: siteIdentifierFromUrl(config!.siteUrl),
        tokenProvider: config!.tokenProvider,
        scopes: config!.scopes,
        graphBaseUrl: config!.graphBaseUrl,
      });
      return client.sites.getByUrl(config!.siteUrl, signal);
    },
  });
}

export function useSiteDrives(
  config: Pick<SharePointConfig, "siteId" | "tokenProvider" | "scopes" | "graphBaseUrl"> | undefined,
) {
  return useQuery({
    queryKey: queryKeys.drives(config?.siteId ?? ""),
    enabled: Boolean(config?.siteId && config?.tokenProvider),
    queryFn: async ({ signal }): Promise<DriveInfo[]> => {
      const client = new SharePointClient({
        siteId: config!.siteId,
        tokenProvider: config!.tokenProvider,
        scopes: config!.scopes,
        graphBaseUrl: config!.graphBaseUrl,
      });
      return client.drives.listDrives(signal);
    },
  });
}

export function useSiteLists(
  config: Pick<SharePointConfig, "siteId" | "tokenProvider" | "scopes" | "graphBaseUrl"> | undefined,
) {
  return useQuery({
    queryKey: queryKeys.lists(config?.siteId ?? ""),
    enabled: Boolean(config?.siteId && config?.tokenProvider),
    queryFn: async ({ signal }): Promise<SharePointListInfo[]> => {
      const client = new SharePointClient({
        siteId: config!.siteId,
        tokenProvider: config!.tokenProvider,
        scopes: config!.scopes,
        graphBaseUrl: config!.graphBaseUrl,
      });
      return client.sites.listLists(config!.siteId, signal);
    },
  });
}

export { findDriveByName, findListByName };

export function usePeopleSearch(query: string, enabled = true) {
  const { client } = useSharePoint();
  const trimmed = query.trim();
  const isTypedSearch = trimmed.length >= 1;
  return useQuery({
    queryKey: queryKeys.people(isTypedSearch ? trimmed : "__suggestions__"),
    enabled,
    queryFn: ({ signal }) =>
      isTypedSearch ? client.people.search(trimmed, signal) : client.people.suggest(signal),
    staleTime: isTypedSearch ? 30_000 : 5 * 60_000,
    retry: 1,
  });
}

export function useFolderChildren(folderId: string | undefined) {
  const { client } = useSharePoint();
  return useQuery({
    queryKey: queryKeys.children(client.config.siteId, client.cacheScope, folderId ?? ""),
    enabled: Boolean(folderId),
    queryFn: ({ signal }) => client.folders.listChildren(folderId!, { signal }),
  });
}

export function useItem(itemId: string | undefined) {
  const { client } = useSharePoint();
  return useQuery({
    queryKey: queryKeys.item(client.config.siteId, client.cacheScope, itemId ?? ""),
    enabled: Boolean(itemId),
    queryFn: ({ signal }) => client.folders.get(itemId!, signal),
  });
}

export function useFolderChildrenInfinite(folderId: string | undefined, expandListItem = false) {
  const { client } = useSharePoint();
  return useInfiniteQuery({
    queryKey: queryKeys.childrenInfinite(
      client.config.siteId,
      client.cacheScope,
      folderId ?? "",
      expandListItem,
    ),
    enabled: Boolean(folderId) && client.config.features.infiniteScroll,
    initialPageParam: undefined as string | undefined,
    queryFn: ({ pageParam, signal }) =>
      client.folders.listChildrenPage(folderId!, {
        nextLink: pageParam,
        expandListItem,
        signal,
      }),
    getNextPageParam: (lastPage) => lastPage.nextLink,
  });
}

export function useSearchItems(
  query: string,
  options: {
    folderId?: string;
    scope?: SearchScope;
    filters?: SearchFilters;
  } = {},
) {
  const { client } = useSharePoint();
  const trimmed = query.trim();
  const scope = options.scope ?? "folder";
  const canSearch =
    trimmed.length >= 2 &&
    client.config.features.search &&
    (scope === "folder" || client.config.features.globalSearch);
  const filterKey = JSON.stringify(options.filters ?? {});

  return useQuery({
    queryKey: queryKeys.search(
      client.config.siteId,
      client.cacheScope,
      `${scope}:${options.folderId ?? ""}:${trimmed}:${filterKey}`,
    ),
    enabled: canSearch,
    queryFn: ({ signal }) =>
      client.search.search({
        query: trimmed,
        folderId: options.folderId,
        scope,
        filters: options.filters,
        signal,
      }),
  });
}

export function useListColumns() {
  const { client } = useSharePoint();
  return useQuery({
    queryKey: queryKeys.listColumns(client.config.siteId, client.cacheScope),
    enabled: client.config.features.metadata,
    queryFn: ({ signal }) => client.listItems.listColumns(signal),
    staleTime: 5 * 60_000,
  });
}

export function useListItemFields(itemId: string | undefined, enabled = true) {
  const { client } = useSharePoint();
  return useQuery({
    queryKey: queryKeys.listItemFields(client.config.siteId, client.cacheScope, itemId ?? ""),
    enabled: Boolean(itemId) && enabled && client.config.features.metadata,
    queryFn: ({ signal }) => client.listItems.getFields(itemId!, signal),
  });
}

export function useUpdateListItemFields(itemId: string) {
  const { client } = useSharePoint();
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (fields: Record<string, string | number | boolean | null>) =>
      client.listItems.updateFields(itemId, fields),
    onSuccess: () => {
      queryClient.invalidateQueries({
        queryKey: queryKeys.listItemFields(client.config.siteId, client.cacheScope, itemId),
      });
      queryClient.invalidateQueries({ queryKey: ["sp", client.config.siteId, client.cacheScope, "children"] });
      queryClient.invalidateQueries({
        queryKey: ["sp", client.config.siteId, client.cacheScope, "children-infinite"],
      });
    },
  });
}

export function useBulkUpdateListItemFields() {
  const { client } = useSharePoint();
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (input: {
      itemIds: string[];
      fields: Record<string, string | number | boolean | null>;
    }) => {
      for (const itemId of input.itemIds) {
        await client.listItems.updateFields(itemId, input.fields);
      }
    },
    onSuccess: (_data, vars) => {
      for (const itemId of vars.itemIds) {
        queryClient.invalidateQueries({
          queryKey: queryKeys.listItemFields(client.config.siteId, client.cacheScope, itemId),
        });
      }
      queryClient.invalidateQueries({ queryKey: ["sp", client.config.siteId, client.cacheScope, "children"] });
      queryClient.invalidateQueries({
        queryKey: ["sp", client.config.siteId, client.cacheScope, "children-infinite"],
      });
    },
  });
}

export function useItemActivities(itemId: string | undefined, enabled = true) {
  const { client } = useSharePoint();
  return useQuery({
    queryKey: queryKeys.activities(client.config.siteId, client.cacheScope, itemId ?? ""),
    enabled: Boolean(itemId) && enabled && client.config.features.activityLog,
    queryFn: ({ signal }) => client.activities.list(itemId!, signal),
  });
}

export { decodeLibraryPage };

export function usePermissions(itemId: string | undefined) {
  const { client } = useSharePoint();
  return useQuery({
    queryKey: queryKeys.permissions(client.config.siteId, client.cacheScope, itemId ?? ""),
    enabled: Boolean(itemId) && client.config.features.manageAccess,
    queryFn: ({ signal }) => client.permissions.list(itemId!, signal),
  });
}

export function useVersions(itemId: string | undefined, enabled: boolean) {
  const { client } = useSharePoint();
  return useQuery({
    queryKey: queryKeys.versions(client.config.siteId, client.cacheScope, itemId ?? ""),
    enabled: Boolean(itemId) && enabled && client.config.features.versionHistory,
    queryFn: ({ signal }) => client.files.getVersions(itemId!, signal),
  });
}

function useInvalidateFolder() {
  const { client } = useSharePoint();
  const queryClient = useQueryClient();
  return (folderId?: string) => {
    if (folderId) {
      return queryClient.invalidateQueries({
        queryKey: queryKeys.children(client.config.siteId, client.cacheScope, folderId),
      });
    }
    return queryClient.invalidateQueries({ queryKey: ["sp", client.config.siteId] });
  };
}

export function useCreateFolder(parentId: string) {
  const { client } = useSharePoint();
  const invalidate = useInvalidateFolder();
  return useMutation({
    mutationFn: (name: string) => client.folders.create(parentId, name),
    onSuccess: () => invalidate(parentId),
  });
}

export function useRenameItem(parentId: string) {
  const { client } = useSharePoint();
  const invalidate = useInvalidateFolder();
  return useMutation({
    mutationFn: ({ itemId, name }: { itemId: string; name: string }) => client.folders.rename(itemId, name),
    onSuccess: () => invalidate(parentId),
  });
}

export function useDeleteItem(parentId: string) {
  const { client } = useSharePoint();
  const invalidate = useInvalidateFolder();
  return useMutation({
    mutationFn: (itemId: string) => client.folders.delete(itemId),
    onSuccess: () => invalidate(parentId),
  });
}

export function useCopyItem(parentId: string) {
  const { client } = useSharePoint();
  const invalidate = useInvalidateFolder();
  return useMutation({
    mutationFn: (input: {
      itemId: string;
      destinationParentId: string;
      newName?: string;
      onCopyProgress?: import("@namphuongso/sharepoint-file-manager-core").CopyMoveOptions["onCopyProgress"];
    }) =>
      client.folders.copy({
        itemId: input.itemId,
        destinationParentId: input.destinationParentId,
        newName: input.newName,
        onCopyProgress: input.onCopyProgress,
      }),
    onSuccess: (_data, vars) => {
      void invalidate(parentId);
      void invalidate(vars.destinationParentId);
    },
  });
}

export function useMoveItem(parentId: string) {
  const { client } = useSharePoint();
  const invalidate = useInvalidateFolder();
  return useMutation({
    mutationFn: (input: { itemId: string; destinationParentId: string; newName?: string }) =>
      client.folders.move(input),
    onSuccess: (_data, vars) => {
      void invalidate(parentId);
      void invalidate(vars.destinationParentId);
    },
  });
}

export function useUploadFile(parentId: string) {
  const { client } = useSharePoint();
  const invalidate = useInvalidateFolder();
  return useMutation({
    mutationFn: (input: {
      file: File;
      onProgress?: (progress: UploadProgress) => void;
      signal?: AbortSignal;
      conflictBehavior?: import("@namphuongso/sharepoint-file-manager-core").ConflictBehavior;
    }) =>
      client.upload.upload({
        parentId,
        fileName: input.file.name,
        content: input.file,
        onProgress: input.onProgress,
        signal: input.signal,
        conflictBehavior: input.conflictBehavior,
      }),
    onSuccess: () => invalidate(parentId),
  });
}

export function useDownloadFile() {
  const { client } = useSharePoint();
  return useMutation({
    mutationFn: async (item: SharePointItem) => {
      await triggerBrowserDownload(client.files.download(item.id));
    },
  });
}

export function useDownloadVersion() {
  const { client } = useSharePoint();
  return useMutation({
    mutationFn: async (input: { itemId: string; versionId: string; fileName?: string }) => {
      const result = await client.files.downloadVersion(input.itemId, input.versionId);
      await triggerBrowserDownload(Promise.resolve(result));
    },
  });
}

async function triggerBrowserDownload(
  resultPromise: Promise<{ blob: Blob; fileName: string; mimeType?: string }>,
) {
  const result = await resultPromise;
  const url = URL.createObjectURL(result.blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = result.fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

export function useInvite(itemId: string) {
  const { client } = useSharePoint();
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (options: InviteOptions) => client.sharing.invite(itemId, options),
    onSuccess: () =>
      queryClient.invalidateQueries({
        queryKey: queryKeys.permissions(client.config.siteId, client.cacheScope, itemId),
      }),
  });
}

export function useCreateLink(itemId: string) {
  const { client } = useSharePoint();
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (options: CreateLinkOptions) => client.sharing.createLink(itemId, options),
    onSuccess: () =>
      queryClient.invalidateQueries({
        queryKey: queryKeys.permissions(client.config.siteId, client.cacheScope, itemId),
      }),
  });
}

export function useRemovePermission(itemId: string) {
  const { client } = useSharePoint();
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (permissionId: string) => client.permissions.remove(itemId, permissionId),
    onSuccess: () =>
      queryClient.invalidateQueries({
        queryKey: queryKeys.permissions(client.config.siteId, client.cacheScope, itemId),
      }),
  });
}

export function useUpdatePermission(itemId: string) {
  const { client } = useSharePoint();
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (input: { permissionId: string; roles: string[] }) =>
      client.permissions.update(itemId, input.permissionId, input.roles),
    onSuccess: () =>
      queryClient.invalidateQueries({
        queryKey: queryKeys.permissions(client.config.siteId, client.cacheScope, itemId),
      }),
  });
}

export function useRestoreVersion(parentId: string) {
  const { client } = useSharePoint();
  const invalidate = useInvalidateFolder();
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (input: { itemId: string; versionId: string }) =>
      client.files.restoreVersion(input.itemId, input.versionId),
    onSuccess: (_data, vars) => {
      void invalidate(parentId);
      void queryClient.invalidateQueries({
        queryKey: queryKeys.versions(client.config.siteId, client.cacheScope, vars.itemId),
      });
    },
  });
}

export function useCheckout(parentId: string) {
  const { client } = useSharePoint();
  const invalidate = useInvalidateFolder();
  return useMutation({
    mutationFn: (input: { itemId: string; action: "checkout" | "checkin" | "discardCheckout"; comment?: string }) => {
      if (input.action === "checkout") return client.checkout.checkout(input.itemId);
      if (input.action === "checkin") return client.checkout.checkin(input.itemId, input.comment);
      return client.checkout.discardCheckout(input.itemId);
    },
    onSuccess: () => invalidate(parentId),
  });
}

export function useCreateOfficeFile(parentId: string) {
  const { client } = useSharePoint();
  const invalidate = useInvalidateFolder();
  return useMutation({
    mutationFn: (kind: import("@namphuongso/sharepoint-file-manager-core").OfficeFileKind) =>
      client.createOfficeFile(parentId, kind),
    onSuccess: () => invalidate(parentId),
  });
}

export function getErrorMessage(
  error: unknown,
  messages: {
    permissionDenied: string;
    notFound: string;
    conflict: string;
    throttled: string;
    authRequired: string;
    unknownError: string;
  },
): string {
  if (!isSharePointError(error)) {
    return error instanceof Error ? error.message : messages.unknownError;
  }
  switch (error.code) {
    case SharePointErrorCode.Forbidden:
      return messages.permissionDenied;
    case SharePointErrorCode.NotFound:
      return messages.notFound;
    case SharePointErrorCode.Conflict:
      return messages.conflict;
    case SharePointErrorCode.Throttled:
      return messages.throttled;
    case SharePointErrorCode.Unauthorized:
    case SharePointErrorCode.InteractionRequired:
      return messages.authRequired;
    default:
      return error.message || messages.unknownError;
  }
}
