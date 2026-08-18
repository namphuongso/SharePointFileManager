import type { ReactNode } from "react";
import type { FeatureConfig, SearchScope, SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import type { OfficeFileKind } from "@namphuongso/sharepoint-file-manager-core";
import {
  Button,
  Menu,
  MenuDivider,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  SearchBox,
  tokens,
} from "@fluentui/react-components";
import {
  AddRegular,
  ArrowClockwiseRegular,
  ArrowDownloadRegular,
  ArrowMoveRegular,
  ArrowUploadRegular,
  ChevronDownRegular,
  ColumnTripleRegular,
  CopyRegular,
  DeleteRegular,
  DismissRegular,
  DocumentOnePageRegular,
  FilterRegular,
  FolderAddRegular,
  GridRegular,
  InfoRegular,
  MoreHorizontalRegular,
  RenameRegular,
  SettingsRegular,
  ShareRegular,
  TextBulletListTreeRegular,
  TextDensityRegular,
} from "@fluentui/react-icons";
import type { Messages } from "../i18n/messages";
import type { SelectionAction } from "./SelectionToolbar";
import { FileTypeChipIcon, type FileKind } from "./FileTypeIcon";

export type FileTypeFilter = Extract<FileKind, "word" | "excel" | "powerpoint" | "pdf">;

const TYPE_FILTERS: FileTypeFilter[] = ["word", "excel", "powerpoint", "pdf"];

export function CommandBar({
  messages,
  features,
  title,
  view,
  selectedItems,
  query,
  searchScope,
  showFilters,
  showColumnChooser,
  detailsOpen,
  typeFilter,
  createOfficePending,
  onViewChange,
  onNewFolder,
  onUpload,
  onUploadFolder,
  onCreateOffice,
  onRefresh,
  onClearSelection,
  onSelectionAction,
  onSearchChange,
  onSearchScopeChange,
  onToggleFilters,
  onToggleColumnChooser,
  onToggleDetails,
  onTypeFilterChange,
}: {
  messages: Messages;
  features: Required<FeatureConfig>;
  title: string;
  view: "list" | "compact" | "grid";
  selectedItems: SharePointItem[];
  query: string;
  searchScope: SearchScope;
  showFilters: boolean;
  showColumnChooser: boolean;
  detailsOpen: boolean;
  typeFilter?: FileTypeFilter;
  createOfficePending?: boolean;
  onViewChange: (view: "list" | "compact" | "grid") => void;
  onNewFolder: () => void;
  onUpload: () => void;
  onUploadFolder: () => void;
  onCreateOffice: (kind: OfficeFileKind) => void;
  onRefresh: () => void;
  onClearSelection: () => void;
  onSelectionAction: (action: SelectionAction) => void;
  onSearchChange: (value: string) => void;
  onSearchScopeChange: (scope: SearchScope) => void;
  onToggleFilters: () => void;
  onToggleColumnChooser: () => void;
  onToggleDetails: () => void;
  onTypeFilterChange: (value?: FileTypeFilter) => void;
}) {
  const hasSelection = selectedItems.length > 0;
  const canCreate = features.createFolder || features.createOfficeFile;
  const createLabel = messages.uploadFiles === "Upload files" ? "Create or upload" : "Tạo hoặc tải lên";

  return (
    <div className="spm-chrome">
      <div className="spm-library-header">
        <h2 className="spm-library-title">{title}</h2>
        <div className="spm-library-header-actions">
          {features.search ? (
            <SearchBox
              className="spm-library-search"
              value={query}
              placeholder={messages.searchLibraryPlaceholder}
              onChange={(_, data) => onSearchChange(data.value)}
            />
          ) : null}
        </div>
      </div>

      <div className="spm-command-surface">
        <div className="spm-command-row">
        <div className="spm-command-links">
          {hasSelection ? (
            <>
              <span className="spm-selection-count">
                {selectedItems.length} {messages.selected}
              </span>
              {features.download && selectedItems.some((item) => item.type === "file") ? (
                <CommandLink icon={<ArrowDownloadRegular />} onClick={() => onSelectionAction("download")}>
                  {messages.download}
                </CommandLink>
              ) : null}
              {features.copy ? (
                <CommandLink icon={<CopyRegular />} onClick={() => onSelectionAction("copy")}>
                  {messages.copy}
                </CommandLink>
              ) : null}
              {features.move ? (
                <CommandLink icon={<ArrowMoveRegular />} onClick={() => onSelectionAction("move")}>
                  {messages.move}
                </CommandLink>
              ) : null}
              {features.share && selectedItems.length === 1 ? (
                <CommandLink icon={<ShareRegular />} onClick={() => onSelectionAction("share")}>
                  {messages.share}
                </CommandLink>
              ) : null}
              <SelectionMoreMenu
                items={selectedItems}
                messages={messages}
                features={features}
                onAction={onSelectionAction}
              />
              {features.delete ? (
                <CommandLink icon={<DeleteRegular />} onClick={() => onSelectionAction("delete")}>
                  {messages.delete}
                </CommandLink>
              ) : null}
              <CommandLink icon={<DismissRegular />} onClick={onClearSelection}>
                {messages.cancel}
              </CommandLink>
            </>
          ) : (
            <>
              {(canCreate || features.upload) ? (
                <CreateMenu
                  label={createLabel}
                  messages={messages}
                  features={features}
                  pending={createOfficePending}
                  onNewFolder={onNewFolder}
                  onUpload={onUpload}
                  onUploadFolder={onUploadFolder}
                  onCreateOffice={onCreateOffice}
                />
              ) : null}
              <CommandLink icon={<ArrowClockwiseRegular />} onClick={onRefresh}>
                {messages.refresh}
              </CommandLink>
              {features.globalSearch && searchScope === "library" ? (
                <CommandLink icon={<FilterRegular />} active={showFilters} onClick={onToggleFilters}>
                  {messages.filters}
                </CommandLink>
              ) : null}
              {features.globalSearch ? (
                <CommandLink
                  onClick={() => onSearchScopeChange(searchScope === "library" ? "folder" : "library")}
                >
                  {searchScope === "library" ? messages.searchScopeLibrary : messages.searchScopeFolder}
                </CommandLink>
              ) : null}
            </>
          )}
        </div>
        </div>

        <div className="spm-view-row">
          <div className="spm-view-tabs">
            <button
              type="button"
              className="spm-view-pill active"
              aria-pressed
            >
              {messages.allDocuments}
            </button>
          </div>

          <div className="spm-type-filters spm-type-filters-compact" role="group" aria-label={messages.fileType}>
            {TYPE_FILTERS.map((kind) => (
              <button
                key={kind}
                type="button"
              className={`spm-type-chip ${typeFilter === kind ? "active" : ""}`}
              title={kind}
              aria-pressed={typeFilter === kind}
                onClick={() => onTypeFilterChange(typeFilter === kind ? undefined : kind)}
              >
                <FileTypeChipIcon kind={kind} size={22} />
              </button>
            ))}
          </div>

          <div className="spm-view-tools">
            {(view === "list" || view === "compact") ? (
              <Button
                appearance={showColumnChooser ? "primary" : "subtle"}
                icon={<ColumnTripleRegular />}
                onClick={onToggleColumnChooser}
                title={messages.columns}
              />
            ) : null}
            <>
              <Button appearance={view === "list" ? "primary" : "subtle"} icon={<TextBulletListTreeRegular />} onClick={() => onViewChange("list")} title={messages.list} />
              <Button appearance={view === "compact" ? "primary" : "subtle"} icon={<TextDensityRegular />} onClick={() => onViewChange("compact")} title={messages.compact} />
              <Button appearance={view === "grid" ? "primary" : "subtle"} icon={<GridRegular />} onClick={() => onViewChange("grid")} title={messages.grid} />
            </>
            {features.properties ? (
              <Button appearance="subtle" icon={<InfoRegular />} onClick={onToggleDetails} className={detailsOpen ? "spm-tool-active" : undefined} title={messages.details}>
                <span className="spm-tool-label">{messages.details}</span>
              </Button>
            ) : null}
          </div>
        </div>
      </div>
    </div>
  );
}

function CreateMenu({
  label,
  messages,
  features,
  pending,
  onNewFolder,
  onUpload,
  onUploadFolder,
  onCreateOffice,
}: {
  label: string;
  messages: Messages;
  features: Required<FeatureConfig>;
  pending?: boolean;
  onNewFolder: () => void;
  onUpload: () => void;
  onUploadFolder: () => void;
  onCreateOffice: (kind: OfficeFileKind) => void;
}) {
  return (
    <Menu>
      <MenuTrigger disableButtonEnhancement>
        <Button appearance="primary" icon={<AddRegular />} className="spm-create-btn">
          {label}
          <ChevronDownRegular />
        </Button>
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          {features.createFolder ? <MenuItem icon={<FolderAddRegular />} onClick={onNewFolder}>{messages.newFolder}</MenuItem> : null}
          {features.createOfficeFile ? (
            <>
              <MenuDivider />
              <MenuItem disabled={pending} onClick={() => onCreateOffice("word")}>{messages.createWord}</MenuItem>
              <MenuItem disabled={pending} onClick={() => onCreateOffice("excel")}>{messages.createExcel}</MenuItem>
              <MenuItem disabled={pending} onClick={() => onCreateOffice("powerpoint")}>{messages.createPowerPoint}</MenuItem>
            </>
          ) : null}
          {features.upload ? (
            <>
              <MenuDivider />
              <MenuItem icon={<ArrowUploadRegular />} onClick={onUpload}>{messages.uploadFiles}</MenuItem>
              <MenuItem icon={<FolderAddRegular />} onClick={onUploadFolder}>{messages.uploadFolder}</MenuItem>
            </>
          ) : null}
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}

function SelectionMoreMenu({
  items,
  messages,
  features,
  onAction,
}: {
  items: SharePointItem[];
  messages: Messages;
  features: Required<FeatureConfig>;
  onAction: (action: SelectionAction) => void;
}) {
  const single = items.length === 1 ? items[0] : undefined;
  const hasActions = Boolean(
    (single && features.rename) ||
      (single?.type === "file" && features.preview) ||
      (single && features.manageAccess) ||
      (items.length > 1 && features.bulkMetadata && features.metadata),
  );
  if (!hasActions) return null;

  return (
    <Menu>
      <MenuTrigger disableButtonEnhancement>
        <Button appearance="subtle" icon={<MoreHorizontalRegular />} className="spm-toolbar-button">
          {messages.moreActions}
        </Button>
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          {single?.type === "file" && features.preview ? (
            <MenuItem icon={<DocumentOnePageRegular />} onClick={() => onAction("preview")}>{messages.preview}</MenuItem>
          ) : null}
          {single && features.rename ? (
            <MenuItem icon={<RenameRegular />} onClick={() => onAction("rename")}>{messages.rename}</MenuItem>
          ) : null}
          {single && features.manageAccess ? (
            <MenuItem icon={<SettingsRegular />} onClick={() => onAction("manageAccess")}>{messages.manageAccess}</MenuItem>
          ) : null}
          {items.length > 1 && features.bulkMetadata && features.metadata ? (
            <MenuItem icon={<ColumnTripleRegular />} onClick={() => onAction("bulkMetadata")}>{messages.bulkEditMetadata}</MenuItem>
          ) : null}
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}

function CommandLink({
  children,
  icon,
  onClick,
  active,
}: {
  children: string;
  icon?: ReactNode;
  onClick?: () => void;
  active?: boolean;
}) {
  return (
    <button
      type="button"
      className={`spm-command-link ${active ? "active" : ""}`}
      onClick={onClick}
      style={{ color: tokens.colorNeutralForeground1 }}
    >
      {icon ? <span className="spm-command-link-icon">{icon}</span> : null}
      {children}
    </button>
  );
}
