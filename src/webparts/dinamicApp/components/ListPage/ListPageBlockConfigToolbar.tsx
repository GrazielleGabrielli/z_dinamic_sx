import * as React from 'react';
import { ActionButton, Stack, Text } from '@fluentui/react';

export interface IListPageBlockConfigToolbarProps {
  label: string;
  onConfigure: () => void;
}

export const ListPageBlockConfigToolbar: React.FC<IListPageBlockConfigToolbarProps> = ({
  label,
  onConfigure,
}) => (
  <Stack
    horizontal
    horizontalAlign="space-between"
    verticalAlign="center"
    styles={{ root: { marginBottom: 8 } }}
  >
    <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#605e5c' } }}>
      {label}
    </Text>
    <ActionButton
      iconProps={{ iconName: 'Settings' }}
      onClick={onConfigure}
      styles={{ root: { height: 28, color: '#0078d4' } }}
    >
      Configurar
    </ActionButton>
  </Stack>
);
