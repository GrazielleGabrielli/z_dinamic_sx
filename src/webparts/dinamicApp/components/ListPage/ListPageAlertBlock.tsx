import * as React from 'react';
import { useState } from 'react';
import { Link, MessageBar, MessageBarType, Stack, Text, ActionButton } from '@fluentui/react';
import type { IListPageAlertBlockConfig, TListPageAlertVariant } from '../../core/config/types';
import { defaultAlertConfig } from '../../core/listPage/listPageBlockConfigUtils';

export interface IListPageAlertBlockProps {
  alert?: IListPageAlertBlockConfig;
  onConfigure?: () => void;
}

function messageBarTypeForVariant(v: TListPageAlertVariant): MessageBarType {
  if (v === 'success') return MessageBarType.success;
  if (v === 'warning') return MessageBarType.warning;
  if (v === 'error') return MessageBarType.error;
  return MessageBarType.info;
}

const DEFAULT_ICON: Record<TListPageAlertVariant, string> = {
  info: 'Info',
  success: 'CheckMark',
  warning: 'Warning',
  error: 'ErrorBadge',
};

export const ListPageAlertBlock: React.FC<IListPageAlertBlockProps> = ({ alert: raw, onConfigure }) => {
  const c = raw ?? defaultAlertConfig();
  const [dismissed, setDismissed] = useState(false);
  const iconOverride = c.iconName.trim();
  const hasLink = Boolean(c.linkUrl.trim() && c.linkText.trim());
  const hasBody = Boolean(c.title.trim() || c.message.trim() || hasLink);
  if (!hasBody && onConfigure === undefined) {
    return null;
  }

  const toolbar =
    onConfigure !== undefined ? (
      <Stack
        horizontal
        horizontalAlign="space-between"
        verticalAlign="center"
        styles={{ root: { marginBottom: 8 } }}
      >
        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#605e5c' } }}>
          Alerta
        </Text>
        <ActionButton
          iconProps={{ iconName: 'Settings' }}
          onClick={onConfigure}
          styles={{ root: { height: 28, color: '#0078d4' } }}
        >
          Configurar
        </ActionButton>
      </Stack>
    ) : null;

  if (dismissed && c.dismissible) {
    return toolbar;
  }

  const emphasizedStyles = c.emphasized
    ? {
        root: {
          borderWidth: 1,
          borderStyle: 'solid',
          borderColor: '#8a8886',
          borderRadius: 4,
        },
      }
    : undefined;

  const inner = (
    <Stack tokens={{ childrenGap: 8 }}>
      {c.title.trim() ? (
        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
          {c.title}
        </Text>
      ) : null}
      {c.message.trim() ? (
        <Text variant="small" styles={{ root: { lineHeight: 1.5 } }}>
          {c.message}
        </Text>
      ) : (
        onConfigure !== undefined && (
          <Text variant="small" styles={{ root: { color: '#a19f9d', fontStyle: 'italic' } }}>
            Sem mensagem — configure o bloco
          </Text>
        )
      )}
      {hasLink ? (
        <Link href={c.linkUrl.trim()} target="_blank" rel="noopener noreferrer">
          {c.linkText.trim()}
        </Link>
      ) : null}
    </Stack>
  );

  return (
    <>
      {toolbar}
      <MessageBar
        messageBarType={messageBarTypeForVariant(c.variant)}
        isMultiline
        onDismiss={c.dismissible ? () => setDismissed(true) : undefined}
        dismissButtonAriaLabel="Fechar"
        messageBarIconProps={iconOverride ? { iconName: iconOverride } : { iconName: DEFAULT_ICON[c.variant] }}
        styles={emphasizedStyles}
      >
        {inner}
      </MessageBar>
    </>
  );
};
