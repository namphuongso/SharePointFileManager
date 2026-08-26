import {
  FileTypeIcon as FluentFileTypeIcon,
  FileIconType,
  type FileTypeIconSize as FluentIconSize,
} from "@fluentui/react-icons-file-type";
import type { FileTypeIconProps, FileTypeIconSize } from "../../types";

/** Map size token → px hợp lệ của Fluent CDN (16|20|24|32|40|48|64|96). */
function sizePx(size: FileTypeIconSize): FluentIconSize {
  if (size === "lg") return 48;
  if (size === "sm") return 20;
  return 24;
}

/** Lấy đuôi file (không gồm chấm) để Fluent map icon 365. */
function fileExtension(name: string): string | undefined {
  const i = name.lastIndexOf(".");
  if (i <= 0 || i === name.length - 1) return undefined;
  return name.slice(i + 1).toLowerCase();
}

/** Icon cột Tên: bộ file-type Microsoft 365 từ Fluent CDN. */
export function FileTypeIcon({ item, size = "md" }: FileTypeIconProps) {
  const px = sizePx(size);

  if (item.type === "folder") {
    return <FluentFileTypeIcon type={FileIconType.folder} size={px} className="spm-shrink-0" />;
  }

  return (
    <FluentFileTypeIcon
      extension={fileExtension(item.name)}
      size={px}
      className="spm-shrink-0"
    />
  );
}
