import * as React from 'react';
import { Spinner, SpinnerSize, Stack, Text } from '@fluentui/react';
import { DINAMIC_SX_TABLE_CLASS } from './tableLayoutClasses';

export const TableLoadingState: React.FC = () => (
  <Stack
    className={DINAMIC_SX_TABLE_CLASS.loading}
    horizontalAlign="center"
    verticalAlign="center"
    tokens={{ childrenGap: 12 }}
    styles={{ root: { padding: 48 } }}
  >
    <Spinner size={SpinnerSize.medium} />
    <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
      Carregando...
    </Text>
  </Stack>
);
