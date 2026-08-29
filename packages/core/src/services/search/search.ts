import type { SharePointRestClient } from "../../rest/client";
import { mapSearchRow } from "../../mappers/search-row";
import type { LibraryContext } from "../../types/models";
import type {
  RestSearchQueryResponse,
  SearchAccessibleOptions,
  SearchAccessiblePage,
} from "../../types/search";
import { libraryAbsoluteUrl, searchQueryPath } from "./search-query-path";
import { DEFAULT_SEARCH_PAGE_SIZE, librarySearchQuery } from "./search-query";

/**
 * Tìm file/folder user hiện tại được xem trong thư viện (Search REST).
 * Security trim theo token; phân trang StartRow; sortlist khi có sort.
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/general-development/sharepoint-search-rest-api-overview
 */
export class SearchService {
  constructor(
    private readonly rest: SharePointRestClient,
    private readonly getLibrary: () => Promise<LibraryContext>,
  ) {}

  /** Flat list trong Path thư viện — không thay browse một cấp. */
  async listAccessible(options: SearchAccessibleOptions = {}): Promise<SearchAccessiblePage> {
    const library = await this.getLibrary();
    const rowLimit = options.rowLimit ?? DEFAULT_SEARCH_PAGE_SIZE;
    const startRow = options.startRow ?? 0;
    const fieldInternalNames = options.fieldInternalNames ?? [];
    const absUrl = libraryAbsoluteUrl(this.rest.siteUrl, library.rootFolderServerRelativeUrl);

    const body = await this.rest.get<RestSearchQueryResponse>(searchQueryPath(), {
      query: librarySearchQuery({
        siteUrl: this.rest.siteUrl,
        rootFolderServerRelativeUrl: library.rootFolderServerRelativeUrl,
        rowLimit,
        startRow,
        fieldInternalNames,
        sort: options.sort,
      }),
      signal: options.signal,
    });

    const relevant = body.PrimaryQueryResult?.RelevantResults;
    const rows = relevant?.Table?.Rows ?? [];
    const items = rows
      .map((row) => mapSearchRow(row, absUrl, fieldInternalNames))
      .filter((item): item is NonNullable<typeof item> => item !== undefined);

    const totalRows = relevant?.TotalRows ?? 0;
    const rowCount = relevant?.RowCount ?? rows.length;
    // Trang đủ rowLimit → còn trang sau (TotalRows trên SPO đôi khi không khớp).
    const nextStartRow = rowCount >= rowLimit ? startRow + rowCount : undefined;

    return { items, nextStartRow, totalRows };
  }
}
