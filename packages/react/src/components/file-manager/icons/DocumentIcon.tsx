import type { DocumentIconProps } from "../../../types";

export function DocumentIcon({ color, accent, glyph, size }: DocumentIconProps) {
  const glyphSize = glyph.length > 1 ? size * 0.28 : size * 0.42;
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" fill="none" aria-hidden className="spm-shrink-0">
      <path
        d="M7 3.5h11.2L26.5 12v14.3c0 1.5-1.2 2.7-2.7 2.7H7c-1.5 0-2.7-1.2-2.7-2.7V6.2C4.3 4.7 5.5 3.5 7 3.5Z"
        fill={color}
      />
      <path d="M18.2 3.5v6.2c0 1.2 1 2.2 2.2 2.2h6.1L18.2 3.5Z" fill={accent} />
      <text
        x="14.5"
        y="23.5"
        textAnchor="middle"
        fill="#FFFFFF"
        fontSize={glyphSize}
        fontWeight="700"
        fontFamily='"Segoe UI", sans-serif'
      >
        {glyph}
      </text>
    </svg>
  );
}
