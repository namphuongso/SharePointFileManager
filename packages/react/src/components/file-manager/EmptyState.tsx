import { Text } from "@fluentui/react-components";
import { FolderOpenRegular } from "@fluentui/react-icons";
import type { EmptyStateProps } from "../../types";
import { useFileManagerStyles } from "./useFileManagerStyles";

/** Folder trống: icon Regular nét mảnh, không fill. */

export function EmptyState({ messages, title, hint }: EmptyStateProps) {
  const styles = useFileManagerStyles();
  return (
    <div className={styles.emptyState}>
      <FolderOpenRegular className={styles.emptyIcon} />
      <Text size={400} weight="semibold" block>
        {title ?? messages.empty}
      </Text>
      <Text size={300} block className={styles.emptyHint}>
        {hint ?? messages.emptyHint}
      </Text>
    </div>
  );
}
