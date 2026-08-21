import { useEffect, useState } from "react";
import type { SharePointItem } from "@namphuongso/sharepoint-file-manager-core";
import { useFolderChildren } from "../hooks/hooks";
import type { Messages } from "../i18n/messages";
import { FileTypeIcon } from "./FileTypeIcon";

export interface FolderCrumb {
  id: string;
  name: string;
}

export function FolderPicker({
  open,
  rootId,
  rootName,
  initialCrumbs,
  excludeIds,
  messages,
  onFolderChange,
}: {
  open: boolean;
  rootId: string;
  rootName: string;
  initialCrumbs?: FolderCrumb[];
  excludeIds: string[];
  messages: Messages;
  onFolderChange: (folderId: string) => void;
}) {
  const [crumbs, setCrumbs] = useState<FolderCrumb[]>(
    initialCrumbs?.length ? initialCrumbs : [{ id: rootId, name: rootName }],
  );
  const folderId = crumbs[crumbs.length - 1]?.id ?? rootId;
  const query = useFolderChildren(open ? folderId : undefined);
  const folders = (Array.isArray(query.data?.items) ? query.data.items : []).filter(
    (item: SharePointItem) => item.type === "folder" && !excludeIds.includes(item.id),
  );

  useEffect(() => {
    if (!open) return;
    const next = initialCrumbs?.length ? initialCrumbs : [{ id: rootId, name: rootName }];
    setCrumbs(next);
  }, [open, rootId, rootName, initialCrumbs]);

  useEffect(() => {
    if (open) onFolderChange(folderId);
  }, [folderId, open, onFolderChange]);

  return (
    <div>
      <p className="spm-mb-2 spm-text-xs spm-text-sp-muted">{messages.selectDestination}</p>
      <nav className="spm-mb-3 spm-flex spm-flex-wrap spm-items-center spm-gap-1 spm-text-sm">
        {crumbs.map((crumb, index) => (
          <span key={`${crumb.id}-${index}`} className="spm-flex spm-items-center spm-gap-1">
            {index > 0 ? <span className="spm-text-sp-muted">/</span> : null}
            <button
              type="button"
              className="hover:spm-underline"
              onClick={() => setCrumbs((current) => current.slice(0, index + 1))}
            >
              {crumb.name}
            </button>
          </span>
        ))}
      </nav>

      {query.isLoading ? <p className="spm-text-sm spm-text-sp-muted">{messages.loading}</p> : null}

      {!query.isLoading && folders.length === 0 ? (
        <p className="spm-text-sm spm-text-sp-muted">{messages.noSubfolders}</p>
      ) : null}

      <div className="spm-max-h-64 spm-overflow-auto">
        {folders.map((folder) => (
          <button
            type="button"
            key={folder.id}
            className="spm-flex spm-w-full spm-items-center spm-gap-2 spm-rounded-md spm-px-2 spm-py-2 spm-text-left spm-text-sm hover:spm-bg-slate-50"
            onClick={() => setCrumbs((current) => [...current, { id: folder.id, name: folder.name }])}
          >
            <FileTypeIcon item={folder} size="sm" />
            <span className="spm-truncate">{folder.name}</span>
          </button>
        ))}
      </div>
    </div>
  );
}
