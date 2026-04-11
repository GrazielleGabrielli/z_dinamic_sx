import * as React from 'react';
import { ActionButton, DefaultButton, Stack, Text } from '@fluentui/react';
import type { IListPageButtonItemConfig, IListPageButtonsBlockConfig } from '../../core/config/types';
import { defaultButtonsConfig } from '../../core/listPage/listPageBlockConfigUtils';

export interface IListPageButtonsBlockProps {
  buttons?: IListPageButtonsBlockConfig;
  onConfigure?: () => void;
}

function navigateItem(it: IListPageButtonItemConfig): void {
  if (it.actionKind === 'reload') {
    window.location.reload();
    return;
  }
  const u = (it.url ?? '').trim();
  if (!u) return;
  if (it.openInNewTab === true) {
    window.open(u, '_blank', 'noopener,noreferrer');
  } else {
    window.location.assign(u);
  }
}

export const ListPageButtonsBlock: React.FC<IListPageButtonsBlockProps> = ({
  buttons: raw,
  onConfigure,
}) => {
  const cfg = raw ?? defaultButtonsConfig();
  const items = cfg.items ?? [];
  const toolbar =
    onConfigure !== undefined ? (
      <Stack
        horizontal
        horizontalAlign="space-between"
        verticalAlign="center"
        styles={{ root: { marginBottom: 8 } }}
      >
        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#605e5c' } }}>
          Botões
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

  if (items.length === 0 && onConfigure === undefined) {
    return null;
  }

  return (
    <>
      {toolbar}
      <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="center">
        {items.map((it) => (
          <DefaultButton
            key={it.id}
            text={it.label}
            onClick={() => navigateItem(it)}
            styles={{ root: { height: 32 } }}
          />
        ))}
      </Stack>
    </>
  );
};
