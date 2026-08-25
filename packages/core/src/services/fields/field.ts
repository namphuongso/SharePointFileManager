import type { SharePointRestClient } from "../../rest/client";
import { SharePointError, SharePointErrorCode } from "../../errors/sharepoint-error";
import type { LibraryContext, SharePointField } from "../../types/models";
import type { ListFieldsOptions, RestField, RestViewFields } from "../../types/rest";
import { FIXED_LIBRARY_FIELD_NAMES } from "./item-fields";
import { listFieldsPath } from "./list-fields-path";
import { viewFieldsPath } from "./view-fields-path";

/**
 * Cột option lấy từ default view ∩ /fields (Accept-Language).
 * View lỗi: chỉ 3 cột cố định FileLeafRef, Modified, File_x0020_Size.
 */
export class FieldService {
  private listed?: Promise<SharePointField[]>;
  private listedVersion = 0;
  private readonly excluded = new Set<string>();
  private viewFailed = false;
  private locale: string;

  constructor(
    private readonly rest: SharePointRestClient,
    private readonly getLibrary: () => Promise<LibraryContext>,
    initialLocale?: string,
  ) {
    this.locale = initialLocale || "vi-VN";
  }

  /** Đổi Accept-Language cho lần GET fields kế tiếp và bỏ cache nhãn cũ. */
  setLocale(locale?: string): void {
    const nextLocale = locale || "vi-VN";
    if (this.locale === nextLocale) return;
    this.locale = nextLocale;
    this.listed = undefined;
    this.listedVersion += 1;
  }

  async list(options: ListFieldsOptions = {}): Promise<SharePointField[]> {
    if (!this.listed) {
      const version = this.listedVersion;
      this.listed = this.fetch(options.signal).catch((error: unknown) => {
        if (version === this.listedVersion) this.listed = undefined;
        throw error;
      });
    }
    return this.listed;
  }

  exclude(internalName: string): void {
    this.excluded.add(internalName);
    this.listed = undefined;
  }

  private async fetch(signal?: AbortSignal): Promise<SharePointField[]> {
    const library = await this.getLibrary();

    if (!this.viewFailed) {
      try {
        const view = await this.rest.get<RestViewFields>(viewFieldsPath(library.listId), { signal });
        const viewFields = await this.fieldsByInternalNames(
          library.listId,
          [...new Set([...FIXED_LIBRARY_FIELD_NAMES, ...(view?.Items ?? [])])].filter(
            (name) => !this.excluded.has(name),
          ),
          signal,
        );
        if (viewFields.length > 0) return viewFields;
      } catch (error) {
        this.viewFailed = true;
        if (error instanceof SharePointError && error.code === SharePointErrorCode.Cancelled) {
          throw error;
        }
      }
    }

    // View lỗi / rỗng: chỉ 3 cột cố định — không GET hết schema Hidden eq false.
    return this.fieldsByInternalNames(
      library.listId,
      [...FIXED_LIBRARY_FIELD_NAMES],
      signal,
    );
  }

  private async fieldsByInternalNames(
    listId: string,
    internalNames: readonly string[],
    signal?: AbortSignal,
  ): Promise<SharePointField[]> {
    if (internalNames.length === 0) return [];

    const body = await this.rest.get<{ value?: RestField[] }>(listFieldsPath(listId), {
      query: {
        $select: "Id,InternalName,Title,TypeAsString,Hidden",
        $filter: `(${internalNames
          .map((name) => `InternalName eq '${name.replace(/'/g, "''")}'`)
          .join(" or ")}) and Hidden eq false`,
      },
      headers: this.languageHeaders(),
      signal,
    });

    const fieldsByName = new Map(
      (body.value ?? []).flatMap((field) => {
        const internalName = field.InternalName?.trim();
        const title = field.Title?.trim();
        const typeAsString = field.TypeAsString?.trim();
        return internalName && title && field.Hidden !== true
          ? [[internalName, { internalName, title, typeAsString }] as const]
          : [];
      }),
    );

    return internalNames.flatMap((name) => fieldsByName.get(name) ?? []);
  }

  private languageHeaders(): Record<string, string> | undefined {
    return { "Accept-Language": this.locale };
  }
}
