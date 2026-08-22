import { Text, tokens } from "@fluentui/react-components";
import { DocumentRegular } from "@fluentui/react-icons";
import type { EmptyStateProps } from "../../types";

/** Folder không có file/thư mục con. */

export function EmptyState({ messages }: EmptyStateProps) {
  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        gap: 12,
        padding: 48,
        textAlign: "center",
      }}
    >
      <DocumentRegular style={{ fontSize: 64, color: tokens.colorNeutralForeground4 }} />
      <div>
        <Text size={500} weight="semibold" block>
          {messages.empty}
        </Text>
        <Text block style={{ color: tokens.colorNeutralForeground3, marginTop: 8 }}>
          {messages.emptyHint}
        </Text>
      </div>
    </div>
  );
}
