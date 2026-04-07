import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import {
  Stack,
  Text,
  Spinner,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  Modal,
  DefaultButton,
} from '@fluentui/react';
import type { TFormHistoryPresentationKind } from '../../core/config/types/formManager';
import { ItemsService } from '../../../../services';

export interface IFormItemHistoryUiProps {
  listTitle: string;
  itemId: number;
  presentationKind: TFormHistoryPresentationKind;
  isOpen: boolean;
  onDismiss: () => void;
  title: string;
  subtitle?: string;
}

export const FormItemHistoryUi: React.FC<IFormItemHistoryUiProps> = ({
  listTitle,
  itemId,
  presentationKind,
  isOpen,
  onDismiss,
  title,
  subtitle,
}) => {
  const itemsService = useMemo(() => new ItemsService(), []);
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState<string | undefined>(undefined);
  const [rows, setRows] = useState<
    { versionLabel: string; versionId: number; created?: string; isCurrentVersion?: boolean }[]
  >([]);

  useEffect(() => {
    if (!isOpen || !listTitle.trim() || !itemId) return;
    setLoading(true);
    setErr(undefined);
    itemsService
      .getItemVersions(listTitle.trim(), itemId)
      .then((r) => {
        setRows(r);
        setLoading(false);
      })
      .catch((e) => {
        setErr(e instanceof Error ? e.message : String(e));
        setRows([]);
        setLoading(false);
      });
  }, [isOpen, listTitle, itemId, itemsService]);

  const body = (
    <Stack tokens={{ childrenGap: 12 }}>
      {subtitle && (
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          {subtitle}
        </Text>
      )}
      {err && <MessageBar messageBarType={MessageBarType.error}>{err}</MessageBar>}
      {loading && <Spinner label="A carregar versões…" />}
      {!loading && !err && rows.length === 0 && (
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Nenhuma versão encontrada. Confirme se o controlo de versões está ativo na lista.
        </Text>
      )}
      {!loading &&
        rows.map((r) => (
          <Stack
            key={r.versionId}
            horizontal
            horizontalAlign="space-between"
            verticalAlign="center"
            tokens={{ childrenGap: 8 }}
            styles={{
              root: {
                padding: '8px 10px',
                borderRadius: 4,
                border: '1px solid #edebe9',
                background: r.isCurrentVersion ? '#f3f9ff' : '#faf9f8',
              },
            }}
          >
            <Text styles={{ root: { fontWeight: 600 } }}>{r.versionLabel}</Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              {r.created
                ? new Date(r.created).toLocaleString(undefined, {
                    dateStyle: 'short',
                    timeStyle: 'short',
                  })
                : '—'}
              {r.isCurrentVersion ? ' · atual' : ''}
            </Text>
          </Stack>
        ))}
    </Stack>
  );

  if (presentationKind === 'collapse') {
    if (!isOpen) return null;
    return (
      <Stack
        tokens={{ childrenGap: 12 }}
        styles={{
          root: {
            marginTop: 8,
            padding: 16,
            borderRadius: 4,
            border: '1px solid #edebe9',
            background: '#ffffff',
          },
        }}
      >
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
            {title}
          </Text>
          <DefaultButton text="Fechar" onClick={onDismiss} />
        </Stack>
        {body}
      </Stack>
    );
  }

  if (presentationKind === 'modal') {
    return (
      <Modal isOpen={isOpen} onDismiss={onDismiss} isBlocking>
        <Stack
          tokens={{ childrenGap: 16 }}
          styles={{
            root: {
              margin: '48px auto',
              maxWidth: 560,
              background: '#ffffff',
              padding: 24,
              borderRadius: 4,
              boxShadow: '0 6.4px 14.4px rgba(0,0,0,.13)',
            },
          }}
        >
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
              {title}
            </Text>
            <DefaultButton text="Fechar" onClick={onDismiss} />
          </Stack>
          {body}
        </Stack>
      </Modal>
    );
  }

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.medium}
      headerText={title}
      onDismiss={onDismiss}
      isBlocking
      closeButtonAriaLabel="Fechar"
    >
      {body}
    </Panel>
  );
};
