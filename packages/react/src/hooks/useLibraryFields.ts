import { useQuery } from "@tanstack/react-query";
import { useSharePoint } from "../provider/context";
import { queryKeys } from "./queryKeys";

/** Bộ cột SharePoint trả về. Giữ hook để host đọc cùng nguồn với items. */
export function useLibraryFields() {
  const { client, locale } = useSharePoint();
  return useQuery({
    queryKey: queryKeys.fields(
      client.config.siteId,
      client.cacheScope,
      client.config.locale ?? locale,
    ),
    queryFn: ({ signal }) => client.fields.list({ signal }),
  });
}
