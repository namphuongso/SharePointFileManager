import { useEffect, useMemo, useRef, useState } from "react";
import type { DirectoryPerson } from "@namphuongso/sharepoint-file-manager-core";
import { usePeopleSearch } from "../hooks/hooks";
import type { Messages } from "../i18n/messages";

function initials(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  const first = parts[0]?.[0] ?? "";
  const last = parts.length > 1 ? parts[parts.length - 1]?.[0] ?? "" : "";
  return `${first}${last}`.toUpperCase() || "?";
}

function kindLabel(kind: DirectoryPerson["kind"], messages: Messages): string {
  if (kind === "group") return messages.groups;
  if (kind === "email") return messages.email;
  return messages.people;
}

export function PeoplePicker({
  open,
  selected,
  messages,
  onChange,
}: {
  open: boolean;
  selected: DirectoryPerson[];
  messages: Messages;
  onChange: (people: DirectoryPerson[]) => void;
}) {
  const [query, setQuery] = useState("");
  const [debounced, setDebounced] = useState("");
  const [focused, setFocused] = useState(false);
  const blurTimerRef = useRef<number>();

  useEffect(() => {
    const timer = window.setTimeout(() => setDebounced(query), 300);
    return () => window.clearTimeout(timer);
  }, [query]);

  useEffect(() => {
    if (!open) {
      setQuery("");
      setDebounced("");
      setFocused(false);
    }
  }, [open]);

  useEffect(() => {
    return () => {
      if (blurTimerRef.current) window.clearTimeout(blurTimerRef.current);
    };
  }, []);

  const showDropdown = open && focused;
  const search = usePeopleSearch(debounced, showDropdown);
  const options = useMemo(() => {
    const items = Array.isArray(search.data) ? search.data : [];
    const selectedKeys = new Set(selected.map((person) => person.key));
    return items.filter((person) => !selectedKeys.has(person.key));
  }, [search.data, selected]);

  function selectPerson(person: DirectoryPerson) {
    onChange([...selected, person]);
    setQuery("");
    setDebounced("");
  }

  function removePerson(key: string) {
    onChange(selected.filter((person) => person.key !== key));
  }

  return (
    <div className="spm-space-y-2">
      {selected.length > 0 ? (
        <div className="spm-flex spm-flex-wrap spm-gap-1">
          {selected.map((person) => (
            <span
              key={person.key}
              className="spm-inline-flex spm-items-center spm-gap-1 spm-rounded-full spm-bg-slate-100 spm-px-2 spm-py-1 spm-text-xs"
            >
              {person.displayName}
              <button type="button" className="spm-text-sp-muted" onClick={() => removePerson(person.key)}>
                ×
              </button>
            </span>
          ))}
        </div>
      ) : null}

      <div className="spm-relative">
        <input
          className="spm-w-full spm-rounded-md spm-border spm-border-sp-border spm-px-3 spm-py-2 spm-text-sm"
          placeholder={messages.peopleSearchPlaceholder}
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          onFocus={() => {
            if (blurTimerRef.current) window.clearTimeout(blurTimerRef.current);
            setFocused(true);
          }}
          onBlur={() => {
            blurTimerRef.current = window.setTimeout(() => setFocused(false), 150);
          }}
          onKeyDown={(event) => {
            if (event.key === "Enter") {
              event.preventDefault();
              const first = options[0];
              if (first) selectPerson(first);
            }
            if (event.key === "Escape") {
              setFocused(false);
            }
          }}
        />

        {showDropdown ? (
          <div className="spm-absolute spm-left-0 spm-right-0 spm-z-50 spm-mt-1 spm-max-h-64 spm-overflow-auto spm-rounded-md spm-border spm-border-sp-border spm-bg-white spm-shadow-lg">
            {search.isFetching ? (
              <p className="spm-px-3 spm-py-2 spm-text-sm spm-text-sp-muted">{messages.searchingPeople}</p>
            ) : null}
            {search.isError && !search.isFetching ? (
              <p className="spm-px-3 spm-py-2 spm-text-sm spm-text-red-600">{messages.peopleSearchError}</p>
            ) : null}
            {!search.isFetching && !search.isError && options.length === 0 ? (
              <p className="spm-px-3 spm-py-2 spm-text-sm spm-text-sp-muted">{messages.noResults}</p>
            ) : null}
            {options.map((person) => (
              <button
                type="button"
                key={person.key}
                className="spm-flex spm-w-full spm-items-center spm-gap-2 spm-px-3 spm-py-2 spm-text-left hover:spm-bg-slate-50"
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => selectPerson(person)}
              >
                <span className="spm-flex spm-h-8 spm-w-8 spm-shrink-0 spm-items-center spm-justify-center spm-rounded-full spm-bg-slate-200 spm-text-xs spm-font-semibold">
                  {initials(person.displayName)}
                </span>
                <span className="spm-min-w-0">
                  <span className="spm-block spm-truncate spm-text-sm">{person.displayName}</span>
                  <span className="spm-block spm-truncate spm-text-xs spm-text-sp-muted">
                    {person.email || kindLabel(person.kind, messages)}
                  </span>
                </span>
              </button>
            ))}
          </div>
        ) : null}
      </div>
    </div>
  );
}
