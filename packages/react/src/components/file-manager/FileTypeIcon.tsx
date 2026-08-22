import type { FileTypeIconProps, FileTypeIconSize } from "../../types";
import { DocumentIcon } from "./icons/DocumentIcon";
import { FILE_ICONS } from "./icons/fileIconStyles";
import { FolderIcon } from "./icons/FolderIcon";
import { getFileKind } from "./icons/getFileKind";

function sizePx(size: FileTypeIconSize): number {
  if (size === "lg") return 48;
  if (size === "sm") return 20;
  return 28;
}

/** Icon cột Tên: folder vàng hoặc tài liệu theo loại file. */
export function FileTypeIcon({ item, size = "md" }: FileTypeIconProps) {
  const kind = getFileKind(item);
  const px = sizePx(size);
  if (kind === "folder") return <FolderIcon size={px} />;
  const style = FILE_ICONS[kind];
  return <DocumentIcon color={style.color} accent={style.accent} glyph={style.glyph} size={px} />;
}
