import { Text } from "@fluentui/react-components";
import { LockClosedRegular } from "@fluentui/react-icons";
import type { ForbiddenStateProps } from "../../types";
import { useFileManagerStyles } from "./useFileManagerStyles";

/** Không có ViewListItems trên folder hiện tại. */
export function ForbiddenState({ messages }: ForbiddenStateProps) {
  const styles = useFileManagerStyles();
  return (
    <div className={styles.emptyState}>
      <LockClosedRegular className={styles.emptyIcon} />
      <Text size={400} weight="semibold" block>
        {messages.noViewPermission}
      </Text>
      <Text size={300} block className={styles.emptyHint}>
        {messages.noViewPermissionHint}
      </Text>
    </div>
  );
}
