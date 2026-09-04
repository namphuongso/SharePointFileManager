import type { NewDocumentKind, NewDocumentKindInfo } from "../../types/create-document";
import {
  BLANK_DOCX_BASE64,
  BLANK_PPTX_BASE64,
  BLANK_XLSX_BASE64,
} from "./blank-templates";

export const NEW_DOCUMENT_KINDS: readonly NewDocumentKindInfo[] = [
  { kind: "word", extension: ".docx", defaultBaseName: "Document" },
  { kind: "excel", extension: ".xlsx", defaultBaseName: "Book" },
  { kind: "powerpoint", extension: ".pptx", defaultBaseName: "Presentation" },
] as const;

function bytesFromBase64(base64: string): ArrayBuffer {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

/** Nội dung binary / text cho file trống theo loại. */
export function blankDocumentContent(
  kind: NewDocumentKind,
): { body: Blob; contentType: string } {
  switch (kind) {
    case "word":
      return {
        body: new Blob([bytesFromBase64(BLANK_DOCX_BASE64)], {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }),
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      };
    case "excel":
      return {
        body: new Blob([bytesFromBase64(BLANK_XLSX_BASE64)], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      };
    case "powerpoint":
      return {
        body: new Blob([bytesFromBase64(BLANK_PPTX_BASE64)], {
          type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        }),
        contentType:
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      };
  }
}

export function ensureDocumentFileName(kind: NewDocumentKind, name: string): string {
  const info = NEW_DOCUMENT_KINDS.find((k) => k.kind === kind)!;
  const trimmed = name.trim();
  const lower = trimmed.toLowerCase();
  if (lower.endsWith(info.extension)) return trimmed;
  return `${trimmed}${info.extension}`;
}
