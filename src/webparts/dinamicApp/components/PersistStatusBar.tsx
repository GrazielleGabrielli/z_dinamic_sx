import * as React from 'react';
import { MessageBar, MessageBarType, Spinner, SpinnerSize, Stack } from '@fluentui/react';
import { TPersistStatus } from '../core/persist/types';

interface IPersistStatusBarProps {
  status: TPersistStatus;
}

const CONTAINER_STYLE: React.CSSProperties = {
  position: 'sticky',
  top: 0,
  zIndex: 1000,
};

export const PersistStatusBar: React.FC<IPersistStatusBarProps> = ({ status }) => {
  if (status === 'idle') return null;

  if (status === 'pending') {
    return (
      <div style={CONTAINER_STYLE}>
        <MessageBar messageBarType={MessageBarType.warning} isMultiline={false}>
          Alterações pendentes — salve a página para confirmar
        </MessageBar>
      </div>
    );
  }

  if (status === 'saving') {
    return (
      <div style={CONTAINER_STYLE}>
        <MessageBar messageBarType={MessageBarType.info} isMultiline={false}>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <Spinner size={SpinnerSize.xSmall} />
            <span>Salvando configuração...</span>
          </Stack>
        </MessageBar>
      </div>
    );
  }

  if (status === 'persisting') {
    return (
      <div style={CONTAINER_STYLE}>
        <MessageBar messageBarType={MessageBarType.warning} isMultiline={false}>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <Spinner size={SpinnerSize.xSmall} />
            <span>Atualizando página, aguarde...</span>
          </Stack>
        </MessageBar>
      </div>
    );
  }

  if (status === 'saved') {
    return (
      <div style={CONTAINER_STYLE}>
        <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
          Configuração salva com sucesso
        </MessageBar>
      </div>
    );
  }

  // status === 'error'
  return (
    <div style={CONTAINER_STYLE}>
      <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
        Erro ao salvar configuração — tente novamente
      </MessageBar>
    </div>
  );
};
