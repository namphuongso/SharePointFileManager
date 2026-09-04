import { describe, expect, it } from "vitest";
import { ensureDocumentFileName } from "./document-kinds";

describe("ensureDocumentFileName", () => {
  it("appends extension when missing", () => {
    expect(ensureDocumentFileName("word", "Report")).toBe("Report.docx");
    expect(ensureDocumentFileName("excel", "Book")).toBe("Book.xlsx");
  });

  it("keeps extension when already present", () => {
    expect(ensureDocumentFileName("word", "Report.docx")).toBe("Report.docx");
    expect(ensureDocumentFileName("excel", "Book.XLSX")).toBe("Book.XLSX");
  });
});
