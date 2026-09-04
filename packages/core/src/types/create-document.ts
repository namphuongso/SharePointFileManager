/** Loại file trống tạo qua REST Files/add + template nhúng. */
export type NewDocumentKind = "word" | "excel" | "powerpoint";

export interface CreateDocumentOptions {
  /** Tên leaf (có hoặc không đuôi — service thêm extension nếu thiếu). */
  name: string;
  overwrite?: boolean;
  signal?: AbortSignal;
}

export interface NewDocumentKindInfo {
  kind: NewDocumentKind;
  /** Đuôi file mặc định. */
  extension: string;
  /** Tên mặc định không đuôi (i18n ghi đè ở UI). */
  defaultBaseName: string;
}
