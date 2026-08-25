import { Text, tokens } from "@fluentui/react-components";
import { FolderOpenRegular } from "@fluentui/react-icons";
import type { EmptyStateProps } from "../../types";

/** Folder không có file/thư mục con — empty state kiểu SharePoint. */

export function EmptyState({ messages }: EmptyStateProps) {
  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        gap: 8,
        padding: "72px 24px",
        textAlign: "center",
      }}
    >
      <FolderOpenRegular style={{ fontSize: 48, color: tokens.colorNeutralForeground4 }} />
      <Text size={400} weight="semibold" block>
        {messages.empty}
      </Text>
      <Text size={300} block style={{ color: tokens.colorNeutralForeground3, maxWidth: 360 }}>
        {messages.emptyHint}
      </Text>
    </div>
  );
}
