import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import type { FileKind } from "../../../types";

/** Phân loại icon theo đuôi file. Folder không nhìn extension. */
export function getFileKind(item: SharePointItem): FileKind {
  if (item.type === "folder") return "folder";
  const name = item.name.toLowerCase();
  if (/\.(doc|docx|dotx)$/.test(name)) return "word";
  if (/\.(xls|xlsx|xlsm|csv)$/.test(name)) return "excel";
  if (/\.(ppt|pptx|ppsx)$/.test(name)) return "powerpoint";
  if (name.endsWith(".pdf")) return "pdf";
  if (/\.(png|jpe?g|gif|webp|bmp|svg)$/.test(name)) return "image";
  if (/\.(mp4|mov|avi|wmv|webm)$/.test(name)) return "video";
  if (/\.(zip|rar|7z|tar|gz)$/.test(name)) return "archive";
  return "generic";
}
