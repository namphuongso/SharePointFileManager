/**
 * Đoạn path folder dưới root thư viện (ServerRelativeUrl đã chuẩn hóa, không / cuối).
 * null = folderUrl không nằm trong root.
 */
export function relativeFolderSegments(
  rootServerRelativeUrl: string,
  folderServerRelativeUrl: string,
): string[] | null {
  const root = rootServerRelativeUrl.replace(/\/$/, "") || "/";
  const folder = folderServerRelativeUrl.replace(/\/$/, "") || "/";
  if (folder.toLowerCase() === root.toLowerCase()) return [];
  const prefix = root === "/" ? "/" : `${root}/`;
  if (!folder.toLowerCase().startsWith(prefix.toLowerCase())) return null;
  return folder.slice(prefix.length).split("/").filter(Boolean);
}
