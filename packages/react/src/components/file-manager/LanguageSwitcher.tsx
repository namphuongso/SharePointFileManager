import {
  Button,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
} from "@fluentui/react-components";
import { CheckmarkRegular, GlobeRegular } from "@fluentui/react-icons";
import { useSharePoint } from "../../provider/context";
import { useFileManagerStyles } from "./useFileManagerStyles";

/** Menu VI/EN trên toolbar; SharePoint nhận Accept-Language tương ứng. */
export function LanguageSwitcher() {
  const styles = useFileManagerStyles();
  const { locale, messages, setLocale } = useSharePoint();
  const isVietnamese = locale.toLowerCase().startsWith("vi");

  return (
    <Menu positioning="below-end">
      <MenuTrigger disableButtonEnhancement>
        <Button
          appearance="subtle"
          shape="circular"
          className={styles.commandIconButton}
          icon={<GlobeRegular fontSize={20} />}
          aria-label={messages.language}
          title={messages.language}
        />
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          <MenuItem
            icon={isVietnamese ? <CheckmarkRegular /> : undefined}
            onClick={() => setLocale("vi-VN")}
          >
            Tiếng Việt
          </MenuItem>
          <MenuItem
            icon={!isVietnamese ? <CheckmarkRegular /> : undefined}
            onClick={() => setLocale("en-US")}
          >
            English
          </MenuItem>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}
