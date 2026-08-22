import { useQuery } from "@tanstack/react-query";
import { useSharePoint } from "../provider/context";
import { queryKeys } from "./queryKeys";

/** Schema cột library (option ẩn/hiện). Cache theo site + thư viện. */
export function useLibraryFields() {
  const { client } = useSharePoint();
  return useQuery({
    queryKey: queryKeys.fields(client.config.siteId, client.cacheScope),
    queryFn: ({ signal }) => client.fields.list({ signal }),
  });
}
