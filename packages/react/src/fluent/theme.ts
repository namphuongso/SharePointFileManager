import {
  createLightTheme,
  createDarkTheme,
  type BrandVariants,
  type Theme,
} from "@fluentui/react-components";

/** Microsoft 365 / SharePoint brand ramp. */
const sharePointBrand: BrandVariants = {
  10: "#020305",
  20: "#111723",
  30: "#16263D",
  40: "#193253",
  50: "#1A3F6A",
  60: "#1B4C82",
  70: "#185FA7",
  80: "#0078D4",
  90: "#2899F5",
  100: "#479EF5",
  110: "#62ABF5",
  120: "#77B7F7",
  130: "#96C6FA",
  140: "#B4D6FA",
  150: "#CFE4FA",
  160: "#EBF3FC",
};

/** Fluent 2 radii used by the 2024 SharePoint / OneDrive library. */
export const sharePointLightTheme: Theme = {
  ...createLightTheme(sharePointBrand),
  borderRadiusNone: "0px",
  borderRadiusSmall: "4px",
  borderRadiusMedium: "8px",
  borderRadiusLarge: "12px",
  borderRadiusXLarge: "16px",
  borderRadiusCircular: "10000px",
};

export const sharePointDarkTheme: Theme = {
  ...createDarkTheme(sharePointBrand),
  borderRadiusNone: "0px",
  borderRadiusSmall: "4px",
  borderRadiusMedium: "8px",
  borderRadiusLarge: "12px",
  borderRadiusXLarge: "16px",
  borderRadiusCircular: "10000px",
};
