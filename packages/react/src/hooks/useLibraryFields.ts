import { useQuery } from "@tanstack/react-query";
import { useSharePoint } from "../provider/context";
import { queryKeys } from "./queryKeys";

interface UseLibraryFieldsOptions {
  enabled?: boolean;
}

/** Bộ cột SharePoint trả về. Giữ hook để host đọc cùng nguồn với items. */
export function useLibraryFields(options: UseLibraryFieldsOptions = {}) {
  const { client, locale } = useSharePoint();
  const enabled = options.enabled ?? true;
  return useQuery({
    queryKey: queryKeys.fields(
      client.config.siteId,
      client.cacheScope,
      client.config.locale ?? locale,
    ),
    enabled,
    queryFn: ({ signal }) => client.fields.list({ signal }),
  });
}
