import { Button as FluentButton, MessageBar, MessageBarBody } from "@fluentui/react-components";
import type { ErrorBannerProps } from "../../types";
import { useFileManagerStyles } from "./useFileManagerStyles";

/** Lỗi: nền pill xám nhạt + nút thử lại circular, không thanh đỏ đầy. */
export function ErrorBanner({
  message,
  onRetry,
  retryLabel = "Retry",
}: ErrorBannerProps) {
  const styles = useFileManagerStyles();
  return (
    <MessageBar intent="error" className={styles.errorBanner}>
      <MessageBarBody>
        {message}
        {onRetry ? (
          <FluentButton
            appearance="subtle"
            shape="circular"
            size="small"
            onClick={onRetry}
            className={styles.errorRetry}
          >
            {retryLabel}
          </FluentButton>
        ) : null}
      </MessageBarBody>
    </MessageBar>
  );
}
