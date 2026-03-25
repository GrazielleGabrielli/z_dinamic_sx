import * as React from 'react';
import { MessageBar, MessageBarType, Stack } from '@fluentui/react';
import { DINAMIC_SX_TABLE_CLASS } from './tableLayoutClasses';

export interface ITableErrorStateProps {
  message: string;
}

export const TableErrorState: React.FC<ITableErrorStateProps> = ({ message }) => (
  <Stack className={DINAMIC_SX_TABLE_CLASS.error} styles={{ root: { marginTop: 12 } }}>
    <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
      {message}
    </MessageBar>
  </Stack>
);
