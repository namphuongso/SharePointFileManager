import type { NewDocumentKind } from "@namphuongso/sharepoint-file-manager-core";

/** Hành động menu New / chuột phải trên khung danh sách. */
export type NewItemAction =
  | { type: "folder" }
  | { type: "uploadFiles" }
  | { type: "uploadFolder" }
  | { type: "document"; kind: NewDocumentKind };
