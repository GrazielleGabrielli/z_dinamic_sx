import * as React from 'react';
import { MessageBar, MessageBarType, Stack } from '@fluentui/react';

export interface ITableErrorStateProps {
  message: string;
}

export const TableErrorState: React.FC<ITableErrorStateProps> = ({ message }) => (
  <Stack styles={{ root: { marginTop: 12 } }}>
    <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
      {message}
    </MessageBar>
  </Stack>
);
