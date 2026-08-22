import { Button as FluentButton, MessageBar, MessageBarBody } from "@fluentui/react-components";
import type { ErrorBannerProps } from "../../types";

/** Thanh lỗi trên danh sách + nút thử lại. */
export function ErrorBanner({
  message,
  onRetry,
  retryLabel = "Retry",
}: ErrorBannerProps) {
  return (
    <MessageBar intent="error" style={{ margin: "8px 12px" }}>
      <MessageBarBody>
        {message}
        {onRetry ? (
          <FluentButton appearance="transparent" size="small" onClick={onRetry} style={{ marginLeft: 8 }}>
            {retryLabel}
          </FluentButton>
        ) : null}
      </MessageBarBody>
    </MessageBar>
  );
}
