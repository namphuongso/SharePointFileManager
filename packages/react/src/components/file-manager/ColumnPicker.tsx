import {
  FIXED_LIBRARY_FIELD_NAMES,
  type SharePointField,
} from "@namphuongso/sharepoint-file-manager-core";
import {
  Button,
  Checkbox,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  Text,
} from "@fluentui/react-components";
import { ColumnTripleRegular, LockClosedRegular } from "@fluentui/react-icons";
import { useId, useMemo } from "react";
import { fieldLabel } from "../../i18n/messages";
import type { Messages } from "../../types";
import { useFileManagerStyles } from "./useFileManagerStyles";

export interface ColumnPickerProps {
  fields?: readonly SharePointField[];
  /** Set InternalName các cột đang hiển thị (ngoài cột cố định). */
  visible: Set<string>;
  onVisibleChange: (next: Set<string>) => void;
  label: string;
  messages: Messages;
}

/**
 * Chọn ẩn/hiện cột từ view mặc định. Name / Modified / File Size luôn bật — không dùng checkbox disabled.
 */
export function ColumnPicker({
  fields = [],
  visible,
  onVisibleChange,
  label,
  messages,
}: ColumnPickerProps) {
  const titleId = useId();
  const styles = useFileManagerStyles();
  const { locked, optional } = useMemo(() => {
    const lockedFields: SharePointField[] = [];
    const optionalFields: SharePointField[] = [];
    for (const field of fields) {
      if (FIXED_LIBRARY_FIELD_NAMES.has(field.internalName)) lockedFields.push(field);
      else optionalFields.push(field);
    }
    return { locked: lockedFields, optional: optionalFields };
  }, [fields]);

  function toggle(internalName: string) {
    if (FIXED_LIBRARY_FIELD_NAMES.has(internalName)) return;
    const next = new Set(visible);
    if (next.has(internalName)) next.delete(internalName);
    else next.add(internalName);
    onVisibleChange(next);
  }

  return (
    <Popover positioning={{ position: "below-end", offset: { mainAxis: 8 } }}>
      <PopoverTrigger disableButtonEnhancement>
        <Button
          appearance="subtle"
          shape="circular"
          className={styles.commandIconButton}
          icon={<ColumnTripleRegular fontSize={20} />}
          aria-label={label}
        />
      </PopoverTrigger>
      <PopoverSurface aria-labelledby={titleId} className={styles.pickerSurface}>
        <Text id={titleId} className={styles.pickerHeader}>
          {label}
        </Text>
        <div className={styles.pickerList}>
          {locked.length > 0 ? (
            <>
              <Text className={styles.pickerSectionLabel}>{messages.alwaysVisible}</Text>
              {locked.map((field) => (
                <div key={field.internalName} className={styles.pickerLockedRow}>
                  <LockClosedRegular fontSize={16} className={styles.pickerLockedIcon} />
                  <span>{fieldLabel(messages, field.internalName, field.title)}</span>
                </div>
              ))}
            </>
          ) : null}
          {optional.length > 0 ? (
            <>
              <Text className={styles.pickerSectionLabel}>{messages.moreColumns}</Text>
              {optional.map((field) => (
                <Checkbox
                  key={field.internalName}
                  className={styles.pickerCheckbox}
                  label={fieldLabel(messages, field.internalName, field.title)}
                  checked={visible.has(field.internalName)}
                  onChange={() => toggle(field.internalName)}
                />
              ))}
            </>
          ) : null}
        </div>
      </PopoverSurface>
    </Popover>
  );
}
