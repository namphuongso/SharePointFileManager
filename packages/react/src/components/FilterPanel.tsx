import type { SearchFilters } from "@namphuongso/sharepoint-file-manager-core";
import type { Messages } from "../i18n/messages";
import { Button } from "./ui";

export function FilterPanel({
  open,
  messages,
  filters,
  onChange,
  onApply,
  onClear,
}: {
  open: boolean;
  messages: Messages;
  filters: SearchFilters;
  onChange: (filters: SearchFilters) => void;
  onApply: () => void;
  onClear: () => void;
}) {
  if (!open) return null;

  return (
    <div className="spm-border-b spm-border-sp-border spm-bg-slate-50 spm-px-3 spm-py-3">
      <div className="spm-mb-2 spm-text-xs spm-font-semibold spm-uppercase spm-text-sp-muted">{messages.filters}</div>
      <div className="spm-grid spm-grid-cols-1 spm-gap-2 md:spm-grid-cols-2 lg:spm-grid-cols-4">
        <label className="spm-text-sm">
          <span className="spm-mb-1 spm-block spm-text-sp-muted">{messages.fileType}</span>
          <input
            className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-2 spm-py-1.5"
            placeholder="docx, pdf, xlsx"
            value={filters.fileType ?? ""}
            onChange={(event) => onChange({ ...filters, fileType: event.target.value || undefined })}
          />
        </label>
        <label className="spm-text-sm">
          <span className="spm-mb-1 spm-block spm-text-sp-muted">{messages.modifiedAfter}</span>
          <input
            type="date"
            className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-2 spm-py-1.5"
            value={filters.modifiedAfter ?? ""}
            onChange={(event) => onChange({ ...filters, modifiedAfter: event.target.value || undefined })}
          />
        </label>
        <label className="spm-text-sm">
          <span className="spm-mb-1 spm-block spm-text-sp-muted">{messages.modifiedBefore}</span>
          <input
            type="date"
            className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-2 spm-py-1.5"
            value={filters.modifiedBefore ?? ""}
            onChange={(event) => onChange({ ...filters, modifiedBefore: event.target.value || undefined })}
          />
        </label>
        <label className="spm-text-sm">
          <span className="spm-mb-1 spm-block spm-text-sp-muted">{messages.author}</span>
          <input
            className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-2 spm-py-1.5"
            value={filters.author ?? ""}
            onChange={(event) => onChange({ ...filters, author: event.target.value || undefined })}
          />
        </label>
      </div>
      <div className="spm-mt-3 spm-flex spm-gap-2">
        <Button variant="primary" onClick={onApply}>
          {messages.applyFilters}
        </Button>
        <Button onClick={onClear}>{messages.clearFilters}</Button>
      </div>
    </div>
  );
}
