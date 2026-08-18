import type { ReactNode, SVGProps } from "react";

type IconProps = SVGProps<SVGSVGElement> & { size?: number };

function IconBase({ size = 16, children, ...props }: IconProps & { children: ReactNode }) {
  return (
    <svg width={size} height={size} viewBox="0 0 20 20" fill="currentColor" aria-hidden {...props}>
      {children}
    </svg>
  );
}

export function IconAdd({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M10 2a.75.75 0 0 1 .75.75V9h6.25a.75.75 0 0 1 0 1.5H10.75v6.25a.75.75 0 0 1-1.5 0V10.5H3a.75.75 0 0 1 0-1.5h6.25V2.75A.75.75 0 0 1 10 2Z" />
    </IconBase>
  );
}

export function IconUpload({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M10 2.5a.75.75 0 0 1 .75.75v7.69l2.22-2.22a.75.75 0 1 1 1.06 1.06l-3.5 3.5a.75.75 0 0 1-1.06 0l-3.5-3.5a.75.75 0 1 1 1.06-1.06l2.22 2.22V3.25A.75.75 0 0 1 10 2.5ZM4 14.25a.75.75 0 0 0 0 1.5h12a.75.75 0 0 0 0-1.5H4Z" />
    </IconBase>
  );
}

export function IconDownload({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M10 17.5a.75.75 0 0 1-.53-.22l-3.5-3.5a.75.75 0 1 1 1.06-1.06l2.22 2.22V3.25a.75.75 0 0 1 1.5 0v11.69l2.22-2.22a.75.75 0 1 1 1.06 1.06l-3.5 3.5a.75.75 0 0 1-.53.22Z" />
    </IconBase>
  );
}

export function IconShare({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M13.5 3a2.25 2.25 0 1 0 0 4.5 2.25 2.25 0 0 0 0-4.5ZM11.25 5.25a3.75 3.75 0 1 1 5.3 3.4l-4.2 2.52a3.73 3.73 0 0 1 0 1.66l4.2 2.52a3.75 3.75 0 1 1-.9 1.5l-4.2-2.52a3.75 3.75 0 1 1 0-5.08l4.2-2.52a3.73 3.73 0 0 1-1.4-.5Z" />
    </IconBase>
  );
}

export function IconCopy({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M6 4.5A1.5 1.5 0 0 1 7.5 3h7A1.5 1.5 0 0 1 16 4.5v9A1.5 1.5 0 0 1 14.5 15h-7A1.5 1.5 0 0 1 6 13.5v-9ZM7.5 4.5v9h7v-9h-7ZM4 7.5A1.5 1.5 0 0 0 2.5 9v6.5A1.5 1.5 0 0 0 4 17h7.5a1.5 1.5 0 0 0 1.5-1.5H4V7.5Z" />
    </IconBase>
  );
}

export function IconMove({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M3.5 6.5a.75.75 0 0 1 .75-.75h11.5a.75.75 0 0 1 0 1.5H4.25a.75.75 0 0 1-.75-.75Zm0 4a.75.75 0 0 1 .75-.75h11.5a.75.75 0 0 1 0 1.5H4.25a.75.75 0 0 1-.75-.75Zm0 4a.75.75 0 0 1 .75-.75h7.5a.75.75 0 0 1 0 1.5H4.25a.75.75 0 0 1-.75-.75Z" />
    </IconBase>
  );
}

export function IconDelete({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M8.5 4h3a1.5 1.5 0 0 1 1.5 1.5V7h3.25a.75.75 0 0 1 0 1.5H16v8.25A1.75 1.75 0 0 1 14.25 18H5.75A1.75 1.75 0 0 1 4 16.25V8.5H1.25a.75.75 0 0 1 0-1.5H4V5.5A1.5 1.5 0 0 1 5.5 4Zm4.5 1.5H7V7h6V5.5ZM5.5 8.5v7.75c0 .14.11.25.25.25h8.5a.25.25 0 0 0 .25-.25V8.5h-9Z" />
    </IconBase>
  );
}

export function IconRefresh({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M10 3a7 7 0 0 1 6.32 4H14.5a.75.75 0 0 0-.65 1.13l2.5 4.33a.75.75 0 0 0 1.3 0l2.5-4.33A.75.75 0 0 0 19.32 7H18a8.5 8.5 0 1 0 .75 3.5.75.75 0 0 0-1.5 0A7 7 0 1 1 10 3Z" />
    </IconBase>
  );
}

export function IconRecycleBin({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M8.5 4h3l.5 1.5h4.25a.75.75 0 0 1 0 1.5H3.75a.75.75 0 0 1 0-1.5H7L7.5 4Zm-1 4.5a.75.75 0 0 1 .75.75V14a1 1 0 0 0 1 1h.5a1 1 0 0 0 1-1V9.25a.75.75 0 0 1 1.5 0V14a2.5 2.5 0 0 1-2.5 2.5h-.5A2.5 2.5 0 0 1 6.25 14V9.25a.75.75 0 0 1 .75-.75Zm4 0a.75.75 0 0 1 .75.75V14a1 1 0 0 0 1 1h.5a1 1 0 0 0 1-1V9.25a.75.75 0 0 1 1.5 0V14a2.5 2.5 0 0 1-2.5 2.5h-.5A2.5 2.5 0 0 1 10.25 14V9.25a.75.75 0 0 1 .75-.75Z" />
    </IconBase>
  );
}

export function IconSearch({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M8.5 3a5.5 5.5 0 1 1 0 11 5.5 5.5 0 0 1 0-11Zm0 1.5a4 4 0 1 0 2.63 7.05l3.02 3.02a.75.75 0 1 0 1.06-1.06l-3.02-3.02A4 4 0 0 0 8.5 4.5Z" />
    </IconBase>
  );
}

export function IconList({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M3 4.75a.75.75 0 0 1 .75-.75h12.5a.75.75 0 0 1 0 1.5H3.75A.75.75 0 0 1 3 4.75Zm0 5a.75.75 0 0 1 .75-.75h12.5a.75.75 0 0 1 0 1.5H3.75A.75.75 0 0 1 3 9.75Zm0 5a.75.75 0 0 1 .75-.75h12.5a.75.75 0 0 1 0 1.5H3.75a.75.75 0 0 1-.75-.75Z" />
    </IconBase>
  );
}

export function IconCompactList({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M3 5.25a.75.75 0 0 1 .75-.75h12.5a.75.75 0 0 1 0 1.5H3.75A.75.75 0 0 1 3 5.25Zm0 3.5a.75.75 0 0 1 .75-.75h12.5a.75.75 0 0 1 0 1.5H3.75A.75.75 0 0 1 3 8.75Zm0 3.5a.75.75 0 0 1 .75-.75h12.5a.75.75 0 0 1 0 1.5H3.75a.75.75 0 0 1-.75-.75Z" />
    </IconBase>
  );
}

export function IconGrid({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M3 4.25A1.25 1.25 0 0 1 4.25 3h2.5A1.25 1.25 0 0 1 8 4.25v2.5A1.25 1.25 0 0 1 6.75 8h-2.5A1.25 1.25 0 0 1 3 6.75v-2.5Zm0 7A1.25 1.25 0 0 1 4.25 10h2.5A1.25 1.25 0 0 1 8 11.25v2.5A1.25 1.25 0 0 1 6.75 15h-2.5A1.25 1.25 0 0 1 3 13.75v-2.5Zm9-7A1.25 1.25 0 0 1 13.25 3h2.5A1.25 1.25 0 0 1 17 4.25v2.5A1.25 1.25 0 0 1 15.75 8h-2.5A1.25 1.25 0 0 1 12 6.75v-2.5Zm0 7A1.25 1.25 0 0 1 13.25 10h2.5A1.25 1.25 0 0 1 17 11.25v2.5A1.25 1.25 0 0 1 15.75 15h-2.5A1.25 1.25 0 0 1 12 13.75v-2.5Z" />
    </IconBase>
  );
}

export function IconColumns({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M3 4.25c0-.69.56-1.25 1.25-1.25h11.5c.69 0 1.25.56 1.25 1.25v11.5c0 .69-.56 1.25-1.25 1.25H4.25A1.25 1.25 0 0 1 3 15.75V4.25Zm2.5 1.5v9h3.5v-9H5.5Zm5 0v9h3.5v-9h-3.5Z" />
    </IconBase>
  );
}

export function IconFilter({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M3.5 4.75a.75.75 0 0 1 .75-.75h11.5a.75.75 0 0 1 .53 1.28l-4.28 4.28v4.44a.75.75 0 0 1-1.06.67l-2-1A.75.75 0 0 1 8 13.75v-3.44L3.72 5.28a.75.75 0 0 1-.22-.53Z" />
    </IconBase>
  );
}

export function IconChevronDown({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M5.22 7.22a.75.75 0 0 1 1.06 0L10 10.94l3.72-3.72a.75.75 0 1 1 1.06 1.06l-4.25 4.25a.75.75 0 0 1-1.06 0L5.22 8.28a.75.75 0 0 1 0-1.06Z" />
    </IconBase>
  );
}

export function IconChevronRight({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M7.22 5.22a.75.75 0 0 1 1.06 0l4.25 4.25a.75.75 0 0 1 0 1.06l-4.25 4.25a.75.75 0 0 1-1.06-1.06L10.94 10 7.22 6.28a.75.75 0 0 1 0-1.06Z" />
    </IconBase>
  );
}

export function IconClose({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M4.09 4.09a.75.75 0 0 1 1.06 0L10 8.94l4.85-4.85a.75.75 0 1 1 1.06 1.06L11.06 10l4.85 4.85a.75.75 0 0 1-1.06 1.06L10 11.06l-4.85 4.85a.75.75 0 0 1-1.06-1.06L8.94 10 4.09 5.15a.75.75 0 0 1 0-1.06Z" />
    </IconBase>
  );
}

export function IconMore({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M4.5 10a1.5 1.5 0 1 1 3 0 1.5 1.5 0 0 1-3 0Zm5 0a1.5 1.5 0 1 1 3 0 1.5 1.5 0 0 1-3 0Zm5 0a1.5 1.5 0 1 1 3 0 1.5 1.5 0 0 1-3 0Z" />
    </IconBase>
  );
}

export function IconFolder({ size = 16, ...props }: IconProps) {
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" fill="none" aria-hidden {...props}>
      <path
        d="M4 8.5C4 7.67 4.67 7 5.5 7H12l2 2h12.5c.83 0 1.5.67 1.5 1.5v13c0 .83-.67 1.5-1.5 1.5h-21C4.67 25 4 24.33 4 23.5V8.5Z"
        fill="#FFB900"
      />
      <path d="M4 10h24v13.5c0 .83-.67 1.5-1.5 1.5h-21C4.67 25 4 24.33 4 23.5V10Z" fill="#FFCA28" />
    </svg>
  );
}

export function IconLock({ size, ...props }: IconProps) {
  return (
    <IconBase size={size} {...props}>
      <path d="M6 8.5V7a4 4 0 1 1 8 0v1.5h.75A1.75 1.75 0 0 1 16.5 10.25v6.5A1.75 1.75 0 0 1 14.75 18.5h-9.5A1.75 1.75 0 0 1 3.5 16.75v-6.5A1.75 1.75 0 0 1 5.25 8.5H6Zm1.5 0h5V7a2.5 2.5 0 0 0-5 0v1.5Z" />
    </IconBase>
  );
}

export function IconDocumentEmpty({ size = 32, className }: { size?: number; className?: string }) {
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" fill="none" aria-hidden className={className}>
      <path d="M6 4h12l6 6v18c0 1.1-.9 2-2 2H6c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2Z" fill="#EDEBE9" />
      <path d="M18 4v6h6" fill="#C8C6C4" />
      <path d="M18 4l6 6h-4c-1.1 0-2-.9-2-2V4Z" fill="#A19F9D" opacity="0.5" />
    </svg>
  );
}
