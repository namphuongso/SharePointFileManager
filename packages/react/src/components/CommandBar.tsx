import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import type { FeatureConfig, SearchScope, SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import { canPerformItemAction } from "@namphuongso/sharepoint-file-manager-core";
import {
  Button,
  Menu,
  MenuDivider,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  SearchBox,
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
  SearchRegular,
  SettingsRegular,
  ShareRegular,
  TextBulletListTreeRegular,
  TextDensityRegular,
} from "@fluentui/react-icons";
import type { Messages } from "../i18n/messages";
import type { SelectionAction } from "../types/selection-action";
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
  onViewChange,
  onNewFolder,
  onUpload,
  onUploadFolder,
  onRefresh,
  onClearSelection,
  onSelectionAction,
  onSearchChange,
  onSearchScopeChange,
  onToggleFilters,
  onToggleColumnChooser,
  onToggleDetails,
  onTypeFilterChange,
  viewModes,
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
  onViewChange: (view: "list" | "compact" | "grid") => void;
  onNewFolder: () => void;
  onUpload: () => void;
  onUploadFolder: () => void;
  onRefresh: () => void;
  onClearSelection: () => void;
  onSelectionAction: (action: SelectionAction) => void;
  onSearchChange: (value: string) => void;
  onSearchScopeChange: (scope: SearchScope) => void;
  onToggleFilters: () => void;
  onToggleColumnChooser: () => void;
  onToggleDetails: () => void;
  onTypeFilterChange: (value?: FileTypeFilter) => void;
  viewModes: Array<"list" | "compact" | "grid">;
}) {
  const hasSelection = selectedItems.length > 0;
  const canCreate = features.createFolder || features.upload;
  const createLabel = messages.createOrUpload;
  const commandAreaRef = useRef<HTMLDivElement>(null);
  const [maxInlineCommands, setMaxInlineCommands] = useState<number>(999);

  const selectionDisabled = useMemo(
    () => ({
      download:
        selectedItems.filter((item) => item.type === "file").length === 0 ||
        selectedItems
          .filter((item) => item.type === "file")
          .some((item) => !canPerformItemAction(item, "download")),
      copy: selectedItems.length === 0,
      move: selectedItems.length === 0,
      share: selectedItems.length !== 1 || !canPerformItemAction(selectedItems[0]!, "share"),
      rename: selectedItems.length !== 1 || !canPerformItemAction(selectedItems[0]!, "rename"),
      delete: selectedItems.length === 0,
      preview: selectedItems.length !== 1 || !canPerformItemAction(selectedItems[0]!, "preview"),
      manageAccess:
        selectedItems.length !== 1 || !canPerformItemAction(selectedItems[0]!, "manageAccess"),
    }),
    [selectedItems],
  );

  const primaryCommands = useMemo(
    () =>
      hasSelection
        ? [
            {
              key: "download",
              label: messages.download,
              icon: <ArrowDownloadRegular />,
              show: features.download,
              disabled: selectionDisabled.download,
              onClick: () => onSelectionAction("download"),
              priority: 1,
            },
            {
              key: "copy",
              label: messages.copy,
              icon: <CopyRegular />,
              show: features.copy,
              disabled: selectionDisabled.copy,
              onClick: () => onSelectionAction("copy"),
              priority: 2,
            },
            {
              key: "move",
              label: messages.move,
              icon: <ArrowMoveRegular />,
              show: features.move,
              disabled: selectionDisabled.move,
              onClick: () => onSelectionAction("move"),
              priority: 3,
            },
            {
              key: "share",
              label: messages.share,
              icon: <ShareRegular />,
              show: features.share,
              disabled: selectionDisabled.share,
              onClick: () => onSelectionAction("share"),
              priority: 4,
            },
            {
              key: "delete",
              label: messages.delete,
              icon: <DeleteRegular />,
              show: features.delete,
              disabled: selectionDisabled.delete,
              onClick: () => onSelectionAction("delete"),
              priority: 5,
            },
            {
              key: "cancel",
              label: messages.cancel,
              icon: <DismissRegular />,
              show: true,
              disabled: false,
              onClick: onClearSelection,
              priority: 6,
            },
          ]
        : [
            {
              key: "refresh",
              label: messages.refresh,
              icon: <ArrowClockwiseRegular />,
              show: true,
              disabled: false,
              onClick: onRefresh,
              priority: 1,
            },
            {
              key: "filters",
              label: messages.filters,
              icon: <FilterRegular />,
              show: features.globalSearch && searchScope === "library",
              disabled: false,
              onClick: onToggleFilters,
              priority: 2,
            },
            {
              key: "scope",
              label: searchScope === "library" ? messages.searchScopeLibrary : messages.searchScopeFolder,
              icon: <SearchRegular />,
              show: features.globalSearch,
              disabled: false,
              onClick: () => onSearchScopeChange(searchScope === "library" ? "folder" : "library"),
              priority: 3,
            },
          ],
    [
      features.copy,
      features.delete,
      features.download,
      features.globalSearch,
      features.move,
      features.share,
      hasSelection,
      messages.cancel,
      messages.copy,
      messages.delete,
      messages.download,
      messages.filters,
      messages.move,
      messages.refresh,
      messages.searchScopeFolder,
      messages.searchScopeLibrary,
      messages.share,
      onClearSelection,
      onRefresh,
      onSearchScopeChange,
      onSelectionAction,
      onToggleFilters,
      searchScope,
      selectionDisabled.copy,
      selectionDisabled.delete,
      selectionDisabled.download,
      selectionDisabled.move,
      selectionDisabled.share,
    ],
  ).filter((command) => command.show);

  useEffect(() => {
    const element = commandAreaRef.current;
    if (!element) return;
    const observer = new ResizeObserver(() => {
      const width = element.clientWidth;
      let available = Math.max(180, width - 24);
      const estimated = [...primaryCommands]
        .sort((a, b) => a.priority - b.priority)
        .map((command) => ({ key: command.key, width: Math.max(84, command.label.length * 7 + 44) }));
      let count = 0;
      for (const command of estimated) {
        if (available - command.width < 44) break;
        available -= command.width;
        count += 1;
      }
      setMaxInlineCommands(Math.max(1, count));
    });
    observer.observe(element);
    return () => observer.disconnect();
  }, [primaryCommands]);

  const orderedCommands = [...primaryCommands].sort((a, b) => a.priority - b.priority);
  const inlineCommands = orderedCommands.slice(0, maxInlineCommands);
  const overflowCommands = orderedCommands.slice(maxInlineCommands);

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
        <div ref={commandAreaRef} className="spm-command-links">
          {hasSelection ? (
            <>
              <span className="spm-selection-count">
                {selectedItems.length} {messages.selected}
              </span>
              {inlineCommands.map((command) => (
                <CommandLink key={command.key} icon={command.icon} onClick={command.onClick} disabled={command.disabled}>
                  {command.label}
                </CommandLink>
              ))}
              {overflowCommands.length > 0 ? (
                <OverflowMenu commands={overflowCommands} label={messages.moreActions} />
              ) : null}
              <SelectionMoreMenu
                items={selectedItems}
                messages={messages}
                features={features}
                onAction={onSelectionAction}
                disabledState={selectionDisabled}
              />
            </>
          ) : (
            <>
              {(canCreate || features.upload) ? (
                <CreateMenu
                  label={createLabel}
                  messages={messages}
                  features={features}
                  onNewFolder={onNewFolder}
                  onUpload={onUpload}
                  onUploadFolder={onUploadFolder}
                />
              ) : null}
              {inlineCommands.map((command) => (
                <CommandLink
                  key={command.key}
                  icon={command.icon}
                  onClick={command.onClick}
                  active={command.key === "filters" && showFilters}
                  disabled={command.disabled}
                >
                  {command.label}
                </CommandLink>
              ))}
              {overflowCommands.length > 0 ? (
                <OverflowMenu commands={overflowCommands} label={messages.moreActions} />
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
              aria-label={messages.allDocuments}
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
                aria-pressed={showColumnChooser}
                title={messages.columns}
              />
            ) : null}
            <>
              <Button
                appearance={view === "list" ? "primary" : "subtle"}
                icon={<TextBulletListTreeRegular />}
                aria-pressed={view === "list"}
                onClick={() => onViewChange("list")}
                title={messages.list}
                disabled={!viewModes.includes("list")}
              />
              <Button
                appearance={view === "compact" ? "primary" : "subtle"}
                icon={<TextDensityRegular />}
                aria-pressed={view === "compact"}
                onClick={() => onViewChange("compact")}
                title={messages.compact}
                disabled={!viewModes.includes("compact")}
              />
              <Button
                appearance={view === "grid" ? "primary" : "subtle"}
                icon={<GridRegular />}
                aria-pressed={view === "grid"}
                onClick={() => onViewChange("grid")}
                title={messages.grid}
                disabled={!viewModes.includes("grid")}
              />
            </>
            {features.properties ? (
              <Button
                appearance="subtle"
                icon={<InfoRegular />}
                aria-pressed={detailsOpen}
                onClick={onToggleDetails}
                className={detailsOpen ? "spm-tool-active" : undefined}
                title={messages.details}
              >
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
  onNewFolder,
  onUpload,
  onUploadFolder,
}: {
  label: string;
  messages: Messages;
  features: Required<FeatureConfig>;
  onNewFolder: () => void;
  onUpload: () => void;
  onUploadFolder: () => void;
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
          {features.upload ? (
            <>
              {features.createFolder ? <MenuDivider /> : null}
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
  disabledState,
}: {
  items: SharePointItem[];
  messages: Messages;
  features: Required<FeatureConfig>;
  onAction: (action: SelectionAction) => void;
  disabledState: Record<string, boolean>;
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
            <MenuItem disabled={disabledState.preview} icon={<DocumentOnePageRegular />} onClick={() => onAction("preview")}>{messages.preview}</MenuItem>
          ) : null}
          {single && features.rename ? (
            <MenuItem disabled={disabledState.rename} icon={<RenameRegular />} onClick={() => onAction("rename")}>{messages.rename}</MenuItem>
          ) : null}
          {single && features.manageAccess ? (
            <MenuItem disabled={disabledState.manageAccess} icon={<SettingsRegular />} onClick={() => onAction("manageAccess")}>{messages.manageAccess}</MenuItem>
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
  disabled,
}: {
  children: string;
  icon?: ReactNode;
  onClick?: () => void;
  active?: boolean;
  disabled?: boolean;
}) {
  return (
    <Button
      type="button"
      appearance="subtle"
      size="medium"
      className={`spm-command-link ${active ? "active" : ""}`}
      aria-pressed={active}
      onClick={onClick}
      disabled={disabled}
    >
      {icon ? <span className="spm-command-link-icon">{icon}</span> : null}
      {children}
    </Button>
  );
}

function OverflowMenu({
  commands,
  label,
}: {
  commands: Array<{ key: string; label: string; icon?: JSX.Element; onClick: () => void; disabled?: boolean }>;
  label: string;
}) {
  return (
    <Menu>
      <MenuTrigger disableButtonEnhancement>
        <Button appearance="subtle" icon={<MoreHorizontalRegular />} className="spm-toolbar-button">
          {label}
        </Button>
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          {commands.map((command) => (
            <MenuItem key={command.key} icon={command.icon} disabled={command.disabled} onClick={command.onClick}>
              {command.label}
            </MenuItem>
          ))}
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}
