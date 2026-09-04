import { useQuery } from "@tanstack/react-query";
import {
  isSharePointError,
  SharePointErrorCode,
} from "@namphuongso/sharepoint-file-manager-core";
import { useSharePoint } from "../provider/context";
import { queryKeys } from "./queryKeys";

/** Kết quả gate xem / thêm trên folder hiện tại. */
export interface FolderViewAccess {
  isLoading: boolean;
  isFetching: boolean;
  isError: boolean;
  error: unknown;
  /** GET quyền xong — có thể quyết định gọi listChildren hay không. */
  isReady: boolean;
  canView: boolean;
  /** AddListItems — tạo folder / upload. */
  canAdd: boolean;
  viewDenied: boolean;
  refetch: () => Promise<unknown>;
}

/**
 * Quyền trên folder hiện tại (UniqueId, alias "root" do core xử lý).
 * 403 hoặc canView=false → không gọi listChildren.
 * canAdd dùng cho nút New folder / Upload.
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
  const canAdd = query.isSuccess && query.data.canAdd;
  const viewDenied = (query.isSuccess && !query.data.canView) || forbiddenError;
  const isReady = canView || viewDenied;

  return {
    isLoading: query.isLoading,
    isFetching: query.isFetching,
    isError: query.isError && !forbiddenError,
    error: query.error,
    isReady,
    canView,
    canAdd,
    viewDenied,
    refetch: query.refetch,
  };
}
