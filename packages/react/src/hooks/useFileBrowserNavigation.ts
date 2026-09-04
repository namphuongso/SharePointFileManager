import { useEffect, useRef, useState } from "react";
import { useSharePoint } from "../provider/context";
import type { BreadcrumbCrumb } from "../types";
import {
  readFileBrowserUrl,
  writeFileBrowserUrl,
  type BrowserViewParam,
} from "./fileBrowserUrl";

export type BrowserView = BrowserViewParam;

/**
 * Đồng bộ tab + folder UniqueId với ?view=&folder= — F5 / back giữ vị trí.
 * Không phụ thuộc react-router.
 */
export function useFileBrowserNavigation(rootId: string, rootName: string) {
  const { client } = useSharePoint();
  const initial = readFileBrowserUrl();
  const [view, setViewState] = useState<BrowserView>(initial.view);
  const [crumbs, setCrumbsState] = useState<BreadcrumbCrumb[]>([
    { id: rootId, name: rootName },
  ]);
  const [locationReady, setLocationReady] = useState(!initial.folderId);
  const rootNameRef = useRef(rootName);
  rootNameRef.current = rootName;
  const viewRef = useRef(view);
  viewRef.current = view;
  const crumbsRef = useRef(crumbs);
  crumbsRef.current = crumbs;
  const locationReadyRef = useRef(locationReady);
  locationReadyRef.current = locationReady;
  /** Tránh hydrate lại khi `client` đổi reference nhưng cùng thư viện. */
  const hydrateKey = `${client.config.siteId}:${client.cacheScope}:${rootId}`;

  const rootIdSet = (): Set<string> =>
    new Set(["root", rootId, rootId.replace(/[{}]/g, "").toLowerCase()]);

  // Hydrate breadcrumb từ ?folder= (đổi site / thư viện / root).
  useEffect(() => {
    const { folderId } = readFileBrowserUrl();
    if (!folderId) {
      setCrumbsState([{ id: rootId, name: rootNameRef.current }]);
      setLocationReady(true);
      return;
    }

    let cancelled = false;
    setLocationReady(false);
    void client.folders
      .resolveBreadcrumb(folderId)
      .then((segments) => {
        if (cancelled) return;
        setCrumbsState([
          { id: rootId, name: rootNameRef.current },
          ...segments.map((s) => ({ id: s.id, name: s.name })),
        ]);
        setLocationReady(true);
      })
      .catch(() => {
        if (cancelled) return;
        setCrumbsState([{ id: rootId, name: rootNameRef.current }]);
        writeFileBrowserUrl(
          { view: readFileBrowserUrl().view, folderId: undefined },
          "replace",
          rootIdSet(),
        );
        setLocationReady(true);
      });

    return () => {
      cancelled = true;
    };
    // client ổn định theo hydrateKey (site + library + root).
    // eslint-disable-next-line react-hooks/exhaustive-deps -- cố ý khóa theo hydrateKey
  }, [hydrateKey, rootId]);

  // Back / Forward trình duyệt.
  useEffect(() => {
    function onPopState() {
      const next = readFileBrowserUrl();
      setViewState(next.view);
      viewRef.current = next.view;
      if (!next.folderId) {
        setCrumbsState([{ id: rootId, name: rootNameRef.current }]);
        setLocationReady(true);
        return;
      }
      setLocationReady(false);
      void client.folders
        .resolveBreadcrumb(next.folderId)
        .then((segments) => {
          setCrumbsState([
            { id: rootId, name: rootNameRef.current },
            ...segments.map((s) => ({ id: s.id, name: s.name })),
          ]);
          setLocationReady(true);
        })
        .catch(() => {
          setCrumbsState([{ id: rootId, name: rootNameRef.current }]);
          setLocationReady(true);
        });
    }

    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
    // eslint-disable-next-line react-hooks/exhaustive-deps -- popstate theo hydrateKey
  }, [hydrateKey, rootId]);

  function syncUrl(nextCrumbs: BreadcrumbCrumb[], nextView: BrowserView, mode: "push" | "replace") {
    // Chưa hydrate xong: không ghi URL — tránh setView/tab xóa ?folder= khi crumbs còn [root].
    if (!locationReadyRef.current && mode === "push") return;
    const leaf = nextCrumbs[nextCrumbs.length - 1]?.id;
    const folderId =
      !leaf || leaf === rootId || leaf === "root" ? undefined : leaf.replace(/[{}]/g, "");
    writeFileBrowserUrl({ view: nextView, folderId }, mode, rootIdSet());
  }

  function setView(next: BrowserView) {
    // Đang ở tab library rồi — không đẩy URL / không đụng breadcrumb (tránh nhảy root).
    if (next === viewRef.current) return;
    setViewState(next);
    viewRef.current = next;
    syncUrl(crumbsRef.current, next, "push");
  }

  function setCrumbs(next: BreadcrumbCrumb[] | ((prev: BreadcrumbCrumb[]) => BreadcrumbCrumb[])) {
    setCrumbsState((prev) => {
      const resolved = typeof next === "function" ? next(prev) : next;
      crumbsRef.current = resolved;
      syncUrl(resolved, viewRef.current, "push");
      return resolved;
    });
  }

  /** Đổi tab + breadcrumb cùng lúc (vd. mở folder từ Search). */
  function navigate(nextCrumbs: BreadcrumbCrumb[], nextView: BrowserView = "library") {
    setViewState(nextView);
    viewRef.current = nextView;
    crumbsRef.current = nextCrumbs;
    setCrumbsState(nextCrumbs);
    setLocationReady(true);
    syncUrl(nextCrumbs, nextView, "push");
  }

  return {
    view,
    setView,
    crumbs,
    setCrumbs,
    navigate,
    locationReady,
  };
}
