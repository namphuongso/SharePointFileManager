import { useEffect } from "react";
import type { FileAction } from "../components/ContextMenu";

export function useKeyboardShortcuts(options: {
  enabled: boolean;
  onSelectAll: () => void;
  onDelete: () => void;
  onRename: () => void;
  onRefresh: () => void;
  onOpen: () => void;
}) {
  useEffect(() => {
    if (!options.enabled) return;

    function onKeyDown(event: KeyboardEvent) {
      const target = event.target as HTMLElement | null;
      const tag = target?.tagName?.toLowerCase();
      if (tag === "input" || tag === "textarea" || tag === "select" || target?.isContentEditable) return;

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "a") {
        event.preventDefault();
        options.onSelectAll();
        return;
      }
      if (event.key === "Delete" || event.key === "Backspace") {
        options.onDelete();
        return;
      }
      if (event.key === "F2") {
        event.preventDefault();
        options.onRename();
        return;
      }
      if (event.key === "F5") {
        event.preventDefault();
        options.onRefresh();
        return;
      }
      if (event.key === "Enter") {
        options.onOpen();
      }
    }

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [options]);
}

export type { FileAction };
