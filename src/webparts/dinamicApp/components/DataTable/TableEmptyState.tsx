import * as React from 'react';
import { Text, Stack } from '@fluentui/react';
import { DINAMIC_SX_TABLE_CLASS } from './tableLayoutClasses';

export interface ITableEmptyStateProps {
  message?: string;
}

export const TableEmptyState: React.FC<ITableEmptyStateProps> = ({ message }) => (
  <Stack
    className={DINAMIC_SX_TABLE_CLASS.empty}
    horizontalAlign="center"
    verticalAlign="center"
    styles={{ root: { padding: 48, background: '#faf9f8', borderRadius: 8 } }}
  >
    <Text variant="medium" styles={{ root: { color: '#605e5c' } }}>
      {message ?? 'Nenhum item encontrado.'}
    </Text>
  </Stack>
);
