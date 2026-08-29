import { useQuery } from "@tanstack/react-query";
import {
  isSharePointError,
  SharePointErrorCode,
} from "@namphuongso/sharepoint-file-manager-core";
import { useSharePoint } from "../provider/context";
import { queryKeys } from "./queryKeys";

/** Kết quả gate xem danh sách: quyền sẵn sàng, được xem, hoặc bị từ chối. */
export interface FolderViewAccess {
  isLoading: boolean;
  isError: boolean;
  error: unknown;
  /** GET quyền xong — có thể quyết định gọi listChildren hay không. */
  isReady: boolean;
  canView: boolean;
  viewDenied: boolean;
  refetch: () => Promise<unknown>;
}

/**
 * ViewListItems trên folder hiện tại (UniqueId, alias "root" do core xử lý).
 * 403 hoặc canView=false → không gọi listChildren.
 */
export function useFolderViewCapabilities(folderId: string | undefined): FolderViewAccess {
  const { client } = useSharePoint();
  const query = useQuery({
    queryKey: queryKeys.folderCapabilities(
      client.config.siteId,
      client.cacheScope,
      folderId ?? "",
    ),
    enabled: Boolean(folderId),
    queryFn: ({ signal }) => client.permissions.getFolderCapabilities(folderId!, { signal }),
    staleTime: 60_000,
  });

  const forbiddenError =
    query.isError &&
    isSharePointError(query.error) &&
    query.error.code === SharePointErrorCode.Forbidden;

  const canView = query.isSuccess && query.data.canView;
  const viewDenied = (query.isSuccess && !query.data.canView) || forbiddenError;
  const isReady = canView || viewDenied;

  return {
    isLoading: query.isLoading,
    isError: query.isError && !forbiddenError,
    error: query.error,
    isReady,
    canView,
    viewDenied,
    refetch: query.refetch,
  };
}
