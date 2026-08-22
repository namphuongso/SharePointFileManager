import type { RestList } from "../../types/rest";

/**
 * So tên cấu hình với Title hiển thị, tên folder trên URL, hoặc EntityTypeName.
 * Title và tên URL thường khác nhau nên phải khớp cả ba.
 */
export function matchesLibraryName(list: RestList, libraryName: string): boolean {
  const needle = libraryName.trim().toLowerCase();
  return (
    list.Title?.trim().toLowerCase() === needle ||
    list.RootFolder?.Name?.trim().toLowerCase() === needle ||
    list.EntityTypeName?.trim().toLowerCase() === needle
  );
}
