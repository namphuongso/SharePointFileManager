import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";

export type FileKind = "folder" | "word" | "excel" | "powerpoint" | "pdf" | "image" | "video" | "archive" | "generic";

export function getFileKind(item: SharePointItem): FileKind {
  if (item.type === "folder") return "folder";
  const name = item.name.toLowerCase();
  const mime = item.mimeType?.toLowerCase() ?? "";
  if (/\.(doc|docx|dotx)$/.test(name) || mime.includes("word")) return "word";
  if (/\.(xls|xlsx|xlsm|csv)$/.test(name) || mime.includes("excel") || mime.includes("spreadsheet")) return "excel";
  if (/\.(ppt|pptx|ppsx)$/.test(name) || mime.includes("powerpoint") || mime.includes("presentation")) return "powerpoint";
  if (name.endsWith(".pdf") || mime.includes("pdf")) return "pdf";
  if (/\.(png|jpe?g|gif|webp|bmp|svg)$/.test(name) || mime.startsWith("image/")) return "image";
  if (/\.(mp4|mov|avi|wmv|webm)$/.test(name) || mime.startsWith("video/")) return "video";
  if (/\.(zip|rar|7z|tar|gz)$/.test(name)) return "archive";
  return "generic";
}

function sizePx(size: "sm" | "md" | "lg"): number {
  if (size === "lg") return 48;
  if (size === "sm") return 20;
  return 28;
}

function FolderIcon({ size }: { size: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" fill="none" aria-hidden className="spm-shrink-0">
      <path
        d="M3.5 9.2C3.5 7.7 4.7 6.5 6.2 6.5h6.1l1.7 2.2h12.3c1.5 0 2.7 1.2 2.7 2.7v12.4c0 1.5-1.2 2.7-2.7 2.7H6.2c-1.5 0-2.7-1.2-2.7-2.7V9.2Z"
        fill="#FFB900"
      />
      <path
        d="M3.5 12h25v11.6c0 1.5-1.2 2.7-2.7 2.7H6.2c-1.5 0-2.7-1.2-2.7-2.7V12Z"
        fill="#FFD335"
      />
      <path d="M3.5 12h25v2.2H3.5V12Z" fill="#EAA300" opacity="0.55" />
    </svg>
  );
}

function DocumentIcon({
  color,
  accent,
  glyph,
  size,
}: {
  color: string;
  accent: string;
  glyph: string;
  size: number;
}) {
  const glyphSize = glyph.length > 1 ? size * 0.28 : size * 0.42;
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" fill="none" aria-hidden className="spm-shrink-0">
      <path
        d="M7 3.5h11.2L26.5 12v14.3c0 1.5-1.2 2.7-2.7 2.7H7c-1.5 0-2.7-1.2-2.7-2.7V6.2C4.3 4.7 5.5 3.5 7 3.5Z"
        fill={color}
      />
      <path d="M18.2 3.5v6.2c0 1.2 1 2.2 2.2 2.2h6.1L18.2 3.5Z" fill={accent} />
      <text
        x="14.5"
        y="23.5"
        textAnchor="middle"
        fill="#FFFFFF"
        fontSize={glyphSize}
        fontWeight="700"
        fontFamily='"Segoe UI", sans-serif'
      >
        {glyph}
      </text>
    </svg>
  );
}

const FILE_ICONS: Record<Exclude<FileKind, "folder">, { color: string; accent: string; glyph: string }> = {
  word: { color: "#185ABD", accent: "#103F91", glyph: "W" },
  excel: { color: "#107C41", accent: "#0B5C2E", glyph: "X" },
  powerpoint: { color: "#C43E1C", accent: "#9B2F14", glyph: "P" },
  pdf: { color: "#D13438", accent: "#A4262C", glyph: "PDF" },
  image: { color: "#8764B8", accent: "#5C2E91", glyph: "IMG" },
  video: { color: "#0078D4", accent: "#005A9E", glyph: "VID" },
  archive: { color: "#CA5010", accent: "#8A3410", glyph: "ZIP" },
  generic: { color: "#605E5C", accent: "#484644", glyph: "" },
};

export function FileTypeIcon({ item, size = "md" }: { item: SharePointItem; size?: "sm" | "md" | "lg" }) {
  const kind = getFileKind(item);
  const px = sizePx(size);
  if (kind === "folder") return <FolderIcon size={px} />;
  const style = FILE_ICONS[kind];
  return <DocumentIcon color={style.color} accent={style.accent} glyph={style.glyph} size={px} />;
}

export function FileTypeChipIcon({ kind, size = 20 }: { kind: Exclude<FileKind, "folder" | "generic">; size?: number }) {
  const style = FILE_ICONS[kind];
  return <DocumentIcon color={style.color} accent={style.accent} glyph={style.glyph} size={size} />;
}
