/** File cần tải vào zip — path tương đối trong archive. */
export interface FolderZipFileEntry {
  uniqueId: string;
  /** Path trong zip, dùng `/` (vd. `sub/a.txt`). */
  relativePath: string;
}

/** Thư mục rỗng — giữ cấu trúc trong zip. */
export interface FolderZipDirEntry {
  /** Path kết thúc bằng `/`. */
  relativePath: string;
}
