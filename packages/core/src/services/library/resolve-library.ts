import type { SharePointRestClient } from "../../rest/client";
import { SharePointError, SharePointErrorCode } from "../../errors/sharepoint-error";
import type { LibraryContext } from "../../types/models";
import type { RestList } from "../../types/rest";
import { matchesLibraryName } from "./match-library-name";
import { toLibraryContext } from "./to-library-context";

/** Cột tối thiểu để biết list + folder gốc. RootFolder là quan hệ nên cần $expand. */
const LIST_SELECT =
  "Id,Title,EntityTypeName,RootFolder/ServerRelativeUrl,RootFolder/UniqueId,RootFolder/Name";

/**
 * Tìm document library theo `libraryName` (Title hoặc tên folder trên URL).
 * Gọi một lần từ SharePointClient.getLibrary() rồi cache.
 */
export async function resolveLibrary(
  rest: SharePointRestClient,
  libraryName: string,
): Promise<LibraryContext> {
  const name = libraryName.trim();
  if (!name) {
    throw new Error(
      "Thiếu tên thư viện. Đặt VITE_SPFM_LIBRARY_NAME (và VITE_SPFM_SITE_URL).",
    );
  }

  const list = (await getByTitle(rest, name)) ?? (await findDocumentLibraryByName(rest, name));
  if (!list) {
    throw new SharePointError({
      code: SharePointErrorCode.NotFound,
      message: `Không tìm thấy thư viện "${name}" trên site ${rest.siteUrl}. Kiểm tra VITE_SPFM_SITE_URL và VITE_SPFM_LIBRARY_NAME.`,
    });
  }

  return toLibraryContext(list, name);
}

/** Nhanh khi Title trùng env. OData: nháy trong Title phải nhân đôi. */
async function getByTitle(
  rest: SharePointRestClient,
  libraryName: string,
): Promise<RestList | undefined> {
  const title = libraryName.replace(/'/g, "''");
  try {
    return await rest.get<RestList>(`web/lists/getbytitle('${title}')`, {
      query: { $expand: "RootFolder", $select: LIST_SELECT },
    });
  } catch {
    return undefined;
  }
}

/** BaseTemplate 101 = Document Library. Dự phòng khi Title ≠ tên trên URL. */
async function findDocumentLibraryByName(
  rest: SharePointRestClient,
  libraryName: string,
): Promise<RestList | undefined> {
  const result = await rest.get<{ value?: RestList[] }>("web/lists", {
    query: {
      $filter: "BaseTemplate eq 101",
      $top: 200,
      $expand: "RootFolder",
      $select: LIST_SELECT,
    },
  });
  return (result.value ?? []).find((item) => matchesLibraryName(item, libraryName));
}
