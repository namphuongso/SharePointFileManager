import { makeStyles, tokens } from "@fluentui/react-components";

/**
 * Style Toast: card góc dưới-phải, nền sáng, viền mảnh, bóng nhẹ.
 * Giống Teams/Outlook: title 14 semibold, subtitle 12, nút đóng circular ở góc trên-phải card.
 */
export const useToastStyles = makeStyles({
  root: {
    minWidth: "288px",
    maxWidth: "360px",
    paddingTop: "12px",
    paddingBottom: "12px",
    paddingLeft: "16px",
    paddingRight: "12px",
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke3}`,
    boxShadow: tokens.shadow16,
    display: "grid",
    gridTemplateColumns: "auto 1fr auto",
    columnGap: "12px",
    alignItems: "start",
  },
  media: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    width: "20px",
    height: "20px",
    marginTop: "2px",
    flexShrink: 0,
  },
  mediaSuccess: {
    color: tokens.colorStatusSuccessForeground1,
  },
  mediaInfo: {
    color: tokens.colorBrandForeground1,
  },
  mediaError: {
    color: tokens.colorStatusDangerForeground1,
  },
  body: {
    minWidth: "0",
    display: "flex",
    flexDirection: "column",
    gap: "2px",
  },
  title: {
    color: tokens.colorNeutralForeground1,
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase300,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  subtitle: {
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  close: {
    marginTop: "-4px",
    marginRight: "-4px",
    color: tokens.colorNeutralForeground2,
    borderRadius: tokens.borderRadiusCircular,
  },
});
