import type { FeatureConfig } from "@namphuongso/sharepoint-file-manager-core";
import { Button, Text, tokens } from "@fluentui/react-components";
import { AddRegular, ArrowUploadRegular, DocumentRegular } from "@fluentui/react-icons";
import type { Messages } from "../i18n/messages";

export function EmptyState({
  messages,
  features,
  isSearch,
  onUpload,
  onNewFolder,
}: {
  messages: Messages;
  features: Required<FeatureConfig>;
  isSearch: boolean;
  onUpload: () => void;
  onNewFolder: () => void;
}) {
  return (
    <div className="spm-empty-state">
      <DocumentRegular style={{ fontSize: 96, color: tokens.colorNeutralForeground4 }} />
      <div>
        <Text size={500} weight="semibold" block>
          {isSearch ? messages.noResults : messages.empty}
        </Text>
        {!isSearch ? (
          <Text block style={{ color: tokens.colorNeutralForeground3, marginTop: 8 }}>
            {messages.emptyHint}
          </Text>
        ) : null}
      </div>
      {!isSearch ? (
        <div className="spm-flex spm-gap-2">
          {features.upload ? (
            <Button appearance="primary" icon={<ArrowUploadRegular />} onClick={onUpload}>
              {messages.upload}
            </Button>
          ) : null}
          {features.createFolder ? (
            <Button icon={<AddRegular />} onClick={onNewFolder}>
              {messages.newFolder}
            </Button>
          ) : null}
        </div>
      ) : null}
    </div>
  );
}
