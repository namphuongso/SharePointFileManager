import type { SharePointRestClient } from "../../rest/client";
import { extractMissingField } from "../../errors/extract-missing-field";
import { mapRestListItem } from "../../mappers/rest-list-item";
import type { FieldService } from "../fields";
import type { FolderChildrenPage, LibraryContext } from "../../types/models";
import type {
  ListChildrenOptions,
  RestListItem,
  RestODataCollection,
} from "../../types/rest";
import { parseODataCollection, resolveODataNextLink } from "../../utils";
import { selectableItemFieldNames } from "../fields/item-fields";
import { listItemsOrderby } from "./list-items-orderby";
import { DEFAULT_LIST_PAGE_SIZE, listItemsPath, listItemsQuery } from "./list-items-query";
import { resolveFolderUrl } from "./resolve-folder-url";

/** Ghost $select: loại cột rồi thử lại, tối đa 3 lần GET. */
const MAX_GHOST_FIELD_ATTEMPTS = 3;

/**
 * Liệt kê một cấp file/folder qua list items (chỉ đọc).
 * Phân trang @odata.nextLink. $select đúng cột SharePoint trả về.
 * Cột bị items báo "does not exist" (ghost theo tenant) thì loại rồi thử lại.
 */
export class FolderService {
  private itemFieldNames: readonly string[] = [];

  constructor(
    private readonly rest: SharePointRestClient,
    private readonly getLibrary: () => Promise<LibraryContext>,
    private readonly fields: FieldService,
  ) {}

  async listChildren(
    folderId: string,
    options: ListChildrenOptions = {},
  ): Promise<FolderChildrenPage> {
    const body = options.nextLink
      ? await this.rest.getUrl<RestODataCollection<RestListItem>>(options.nextLink, {
          signal: options.signal,
        })
      : await this.fetchFirstPage(folderId, options);

    const page = parseODataCollection(body);
    const fieldInternalNames = this.itemFieldNames;
    const items = page.value
      .map((item) => mapRestListItem(item, fieldInternalNames))
      .filter((item): item is NonNullable<typeof item> => item !== undefined);

    return {
      items,
      nextLink: resolveODataNextLink(page.nextLink, this.rest.siteUrl),
    };
  }

  private async fetchFirstPage(folderId: string, options: ListChildrenOptions) {
    const library = await this.getLibrary();
    const [fileDirRef, allFields] = await Promise.all([
      resolveFolderUrl(this.rest, this.getLibrary, folderId, options.signal),
      this.fields.list({ signal: options.signal }),
    ]);
    const top = options.top ?? DEFAULT_LIST_PAGE_SIZE;
    let knownFields = allFields.filter(
      (field) => field.typeAsString?.toLowerCase() !== "computed",
    );
    // Sort có thể bỏ khi $orderby báo thiếu cột (khác ghost $select).
    let sort = options.sort;

    for (let attempt = 1; attempt <= MAX_GHOST_FIELD_ATTEMPTS; attempt++) {
      try {
        const selectedFields = selectableItemFieldNames(
          knownFields.map((field) => field.internalName),
        );
        this.itemFieldNames = selectedFields;
        const sortField = knownFields.find((field) => field.internalName === sort?.field);
        return await this.rest.get<RestODataCollection<RestListItem>>(
          listItemsPath(library.listId),
          {
            query: listItemsQuery(
              fileDirRef,
              top,
              selectedFields,
              listItemsOrderby(sort, sortField?.typeAsString),
            ),
            signal: options.signal,
          },
        );
      } catch (error) {
        const missing = extractMissingField(error);
        if (!missing || attempt >= MAX_GHOST_FIELD_ATTEMPTS) throw error;

        let progressed = false;

        // $orderby path (vd. File/Length → 'File') hoặc đúng cột đang sort: bỏ sort.
        if (sort && (sort.field === missing || missing === "File")) {
          sort = undefined;
          progressed = true;
        }

        const next = knownFields.filter((field) => field.internalName !== missing);
        if (next.length !== knownFields.length) {
          this.fields.exclude(missing);
          knownFields = next;
          progressed = true;
        }

        if (!progressed) throw error;
      }
    }

    throw new Error("listChildren: hết lần thử ghost field");
  }
}

