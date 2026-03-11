import * as React from 'react';
import { Spinner, SpinnerSize, Stack, Text } from '@fluentui/react';

export const TableLoadingState: React.FC = () => (
  <Stack
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
