import type { RestList } from "../../types/rest";

/**
 * So tên cấu hình với Title hiển thị, tên folder trên URL, hoặc EntityTypeName.
 * Title và tên URL thường khác nhau nên phải khớp cả ba.
 */
export function matchesLibraryName(list: RestList, libraryName: string): boolean {
  const needle = normalizeName(libraryName);
  return (
    normalizeName(list.Title) === needle ||
    normalizeName(list.RootFolder?.Name) === needle ||
    normalizeName(list.EntityTypeName) === needle
  );
}

function normalizeName(value?: string): string {
  return value?.trim().toLowerCase() ?? "";
}
