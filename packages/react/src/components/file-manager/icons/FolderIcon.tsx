import type { FolderIconProps } from "../../../types";

export function FolderIcon({ size }: FolderIconProps) {
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" fill="none" aria-hidden className="spm-shrink-0">
      <path
        d="M3.5 9.2C3.5 7.7 4.7 6.5 6.2 6.5h6.1l1.7 2.2h12.3c1.5 0 2.7 1.2 2.7 2.7v12.4c0 1.5-1.2 2.7-2.7 2.7H6.2c-1.5 0-2.7-1.2-2.7-2.7V9.2Z"
        fill="#FFB900"
      />
      <path
        d="M3.5 12h25v11.6c0 1.5-1.2 2.7-2.7 2.7H6.2c-1.5 0-2.7-1.2-2.7-2.7V12Z"
        fill="#FFD335"
      />
      <path d="M3.5 12h25v2.2H3.5V12Z" fill="#EAA300" opacity="0.55" />
    </svg>
  );
}
