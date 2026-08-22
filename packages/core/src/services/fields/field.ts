import type { SharePointRestClient } from "../../rest/client";
import { mapRestField } from "../../mappers/rest-field";
import type { LibraryContext, SharePointField } from "../../types/models";
import type { ListFieldsOptions, RestField } from "../../types/rest";
import { listFieldsPath } from "./list-fields-path";

/**
 * Schema cột document library (chỉ đọc).
 * Không $select: REST trả đủ property SP.Field; mapper lấy phần UI cần.
 * @see https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
 */
export class FieldService {
  constructor(
    private readonly rest: SharePointRestClient,
    private readonly getLibrary: () => Promise<LibraryContext>,
  ) {}

  async list(options: ListFieldsOptions = {}): Promise<SharePointField[]> {
    const library = await this.getLibrary();
    const result = await this.rest.get<{ value?: RestField[] }>(listFieldsPath(library.listId), {
      signal: options.signal,
    });

    return (result.value ?? [])
      .map(mapRestField)
      .filter((field): field is SharePointField => field !== undefined);
  }
}
