import type { FileKind, FileIconStyle } from "../../../types";

/** Màu / chữ trên icon tài liệu (không gồm folder). */
export const FILE_ICONS: Record<Exclude<FileKind, "folder">, FileIconStyle> = {
  word: { color: "#185ABD", accent: "#103F91", glyph: "W" },
  excel: { color: "#107C41", accent: "#0B5C2E", glyph: "X" },
  powerpoint: { color: "#C43E1C", accent: "#9B2F14", glyph: "P" },
  pdf: { color: "#D13438", accent: "#A4262C", glyph: "PDF" },
  image: { color: "#8764B8", accent: "#5C2E91", glyph: "IMG" },
  video: { color: "#0078D4", accent: "#005A9E", glyph: "VID" },
  archive: { color: "#CA5010", accent: "#8A3410", glyph: "ZIP" },
  generic: { color: "#605E5C", accent: "#484644", glyph: "" },
};
