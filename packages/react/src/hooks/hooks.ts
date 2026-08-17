import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import {
  SharePointClient,
  findDriveByName,
  findListByName,
  isSharePointError,
  SharePointErrorCode,
  siteIdentifierFromUrl,
  type CreateLinkOptions,
  type DriveInfo,
  type InviteOptions,
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

export function useSearchItems(query: string, folderId?: string) {
  const { client } = useSharePoint();
  const trimmed = query.trim();
  return useQuery({
    queryKey: queryKeys.search(client.config.siteId, client.cacheScope, `${folderId ?? ""}:${trimmed}`),
    enabled: trimmed.length >= 2 && client.config.features.search,
    queryFn: ({ signal }) => client.search.search({ query: trimmed, folderId, signal }),
  });
}

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
    mutationFn: (input: { itemId: string; destinationParentId: string; newName?: string }) =>
      client.folders.copy(input),
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
    }) =>
      client.upload.upload({
        parentId,
        fileName: input.file.name,
        content: input.file,
        onProgress: input.onProgress,
        signal: input.signal,
      }),
    onSuccess: () => invalidate(parentId),
  });
}

export function useDownloadFile() {
  const { client } = useSharePoint();
  return useMutation({
    mutationFn: async (item: SharePointItem) => {
      const result = await client.files.download(item.id);
      const url = URL.createObjectURL(result.blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = result.fileName;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      URL.revokeObjectURL(url);
    },
  });
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
