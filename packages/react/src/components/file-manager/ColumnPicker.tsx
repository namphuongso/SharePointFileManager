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
  tokens,
} from "@fluentui/react-components";
import { ColumnTripleRegular } from "@fluentui/react-icons";
import { useId } from "react";
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
 * Chọn ẩn/hiện cột từ view mặc định. Name / Modified / File Size luôn bật.
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

  function toggle(internalName: string) {
    if (FIXED_LIBRARY_FIELD_NAMES.has(internalName)) return;
    const next = new Set(visible);
    if (next.has(internalName)) next.delete(internalName);
    else next.add(internalName);
    onVisibleChange(next);
  }

  return (
    <Popover positioning="below-end">
      <PopoverTrigger disableButtonEnhancement>
        <Button
          appearance="subtle"
          shape="circular"
          className={styles.commandIconButton}
          icon={<ColumnTripleRegular fontSize={20} />}
          aria-label={label}
          title={label}
        />
      </PopoverTrigger>
      <PopoverSurface aria-labelledby={titleId} className={styles.pickerSurface}>
        <Text
          id={titleId}
          weight="semibold"
          size={200}
          style={{
            display: "block",
            padding: "0 16px 8px",
            color: tokens.colorNeutralForeground3,
            textTransform: "uppercase",
            letterSpacing: "0.05em",
          }}
        >
          {label}
        </Text>
        {fields.map((field) => {
          const fixed = FIXED_LIBRARY_FIELD_NAMES.has(field.internalName);
          return (
            <div key={field.internalName} style={{ padding: "2px 12px" }}>
              <Checkbox
                label={fieldLabel(messages, field.internalName, field.title)}
                checked={fixed || visible.has(field.internalName)}
                disabled={fixed}
                onChange={() => toggle(field.internalName)}
              />
            </div>
          );
        })}
      </PopoverSurface>
    </Popover>
  );
}
