import type { PointerEvent as ReactPointerEvent } from "react";
import { useRef, useState } from "react";
import { mergeClasses } from "@fluentui/react-components";
import { useFileManagerStyles } from "./useFileManagerStyles";

export interface ColumnResizeHandleProps {
  width: number;
  minWidth: number;
  label: string;
  onResize: (width: number) => void;
  onResizeEnd: (width: number) => void;
}

/** Tay cầm mép cột: kéo ngang đổi width; chặn drag-reorder của header. */
export function ColumnResizeHandle({
  width,
  minWidth,
  label,
  onResize,
  onResizeEnd,
}: ColumnResizeHandleProps) {
  const styles = useFileManagerStyles();
  const [active, setActive] = useState(false);
  const startRef = useRef({ x: 0, width: 0 });
  const liveRef = useRef(width);
  liveRef.current = width;

  function onPointerDown(event: ReactPointerEvent<HTMLButtonElement>) {
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.setPointerCapture(event.pointerId);
    setActive(true);
    startRef.current = { x: event.clientX, width };
  }

  function onPointerMove(event: ReactPointerEvent<HTMLButtonElement>) {
    if (!event.currentTarget.hasPointerCapture(event.pointerId)) return;
    const next = Math.max(minWidth, Math.round(startRef.current.width + event.clientX - startRef.current.x));
    liveRef.current = next;
    onResize(next);
  }

  function onPointerUp(event: ReactPointerEvent<HTMLButtonElement>) {
    if (!event.currentTarget.hasPointerCapture(event.pointerId)) return;
    event.currentTarget.releasePointerCapture(event.pointerId);
    setActive(false);
    onResizeEnd(liveRef.current);
  }

  return (
    <button
      type="button"
      className={mergeClasses(styles.columnResizeHandle, active && styles.columnResizeHandleActive)}
      aria-label={label}
      draggable={false}
      onPointerDown={onPointerDown}
      onPointerMove={onPointerMove}
      onPointerUp={onPointerUp}
      onPointerCancel={onPointerUp}
      onDragStart={(event) => event.preventDefault()}
    />
  );
}
