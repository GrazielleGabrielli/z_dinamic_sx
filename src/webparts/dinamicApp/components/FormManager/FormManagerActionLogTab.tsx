import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import {
  Stack,
  Text,
  Dropdown,
  IDropdownOption,
  Spinner,
  MessageBar,
  MessageBarType,
  Toggle,
} from '@fluentui/react';
import { ListsService, FieldsService } from '../../../../services';
import type { IListSummary, IFieldMetadata } from '../../../../services';
import type { IFormCustomButtonConfig } from '../../core/config/types/formManager';
import { ListPageRichQuillEditor } from '../ListPage/ListPageRichQuillEditor';

const LOG_QUILL_PERMISSIONS = {
  allowHeaders: true,
  allowLists: true,
  allowLinks: true,
  allowImages: false,
  allowVideoEmbed: false,
};

export interface IFormManagerActionLogTabProps {
  captureEnabled: boolean;
  onCaptureEnabledChange: (enabled: boolean) => void;
  listTitle: string;
  onListTitleChange: (title: string) => void;
  actionFieldInternalName: string;
  onActionFieldInternalNameChange: (internalName: string) => void;
  descriptionsHtmlByButtonId: Record<string, string>;
  onDescriptionChange: (buttonId: string, html: string) => void;
  customButtons: IFormCustomButtonConfig[];
}

export function FormManagerActionLogTabContent(props: IFormManagerActionLogTabProps): JSX.Element {
  const {
    captureEnabled,
    onCaptureEnabledChange,
    listTitle,
    onListTitleChange,
    actionFieldInternalName,
    onActionFieldInternalNameChange,
    descriptionsHtmlByButtonId,
    onDescriptionChange,
    customButtons,
  } = props;
  const [lists, setLists] = useState<IListSummary[]>([]);
  const [listsLoading, setListsLoading] = useState(false);
  const [listsErr, setListsErr] = useState<string | undefined>(undefined);
  const [logListFields, setLogListFields] = useState<IFieldMetadata[]>([]);
  const [logFieldsLoading, setLogFieldsLoading] = useState(false);
  const [logFieldsErr, setLogFieldsErr] = useState<string | undefined>(undefined);

  useEffect(() => {
    setListsErr(undefined);
    setListsLoading(true);
    const svc = new ListsService();
    svc
      .getLists(false)
      .then((data) => {
        setLists(data.filter((l) => !l.IsLibrary));
        setListsLoading(false);
      })
      .catch((e) => {
        setLists([]);
        setListsLoading(false);
        setListsErr(e instanceof Error ? e.message : String(e));
      });
  }, []);

  useEffect(() => {
    if (!listTitle.trim()) {
      setLogListFields([]);
      setLogFieldsErr(undefined);
      return;
    }
    setLogFieldsErr(undefined);
    setLogFieldsLoading(true);
    const fs = new FieldsService();
    fs.getVisibleFields(listTitle.trim())
      .then((fields) => {
        setLogListFields(fields);
        setLogFieldsLoading(false);
      })
      .catch((e) => {
        setLogListFields([]);
        setLogFieldsLoading(false);
        setLogFieldsErr(e instanceof Error ? e.message : String(e));
      });
  }, [listTitle]);

  const multilineFields = useMemo(
    () => logListFields.filter((f) => f.MappedType === 'multiline' && !f.Hidden && !f.ReadOnlyField),
    [logListFields]
  );

  const fieldOptions: IDropdownOption[] = useMemo(() => {
    const base: IDropdownOption[] = [{ key: '', text: '—' }];
    const known = new Set(multilineFields.map((f) => f.InternalName));
    if (actionFieldInternalName && !known.has(actionFieldInternalName)) {
      base.push({
        key: actionFieldInternalName,
        text: `${actionFieldInternalName} (referência guardada)`,
      });
    }
    return base.concat(
      multilineFields.map((f) => ({
        key: f.InternalName,
        text: `${f.Title} (${f.InternalName})`,
      }))
    );
  }, [multilineFields, actionFieldInternalName]);

  const listOptions: IDropdownOption[] = useMemo(
    () => [{ key: '', text: '— nenhuma —' }, ...lists.map((l) => ({ key: l.Title, text: l.Title }))],
    [lists]
  );

  const canEnableCapture = Boolean(listTitle.trim() && actionFieldInternalName.trim());

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { marginTop: 12 } }}>
      <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
        Lista de logs
      </Text>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Escolha a lista de registo, depois um campo de texto multilinhas onde será gravado o texto da ação (metadata).
        Só então pode ativar a captação.
      </Text>
      {listsLoading && <Spinner label="A carregar listas…" />}
      {listsErr && <MessageBar messageBarType={MessageBarType.error}>{listsErr}</MessageBar>}
      <Dropdown
        label="Lista para registos de log"
        options={listOptions}
        selectedKey={listTitle || ''}
        onChange={(_, o) => {
          const t = o ? String(o.key) : '';
          onListTitleChange(t);
          onActionFieldInternalNameChange('');
          if (captureEnabled) onCaptureEnabledChange(false);
        }}
        styles={{ root: { maxWidth: 480 } }}
        disabled={listsLoading}
      />
      {listTitle.trim() ? (
        <>
          {logFieldsLoading && <Spinner label="A carregar campos da lista…" />}
          {logFieldsErr && <MessageBar messageBarType={MessageBarType.error}>{logFieldsErr}</MessageBar>}
          {!logFieldsLoading && !logFieldsErr && multilineFields.length === 0 && (
            <MessageBar messageBarType={MessageBarType.warning}>
              Esta lista não tem colunas de texto multilinhas visíveis. Crie uma coluna «Várias linhas de texto» na
              lista e volte a abrir o painel.
            </MessageBar>
          )}
          <Dropdown
            label="Campo para guardar a ação (só várias linhas)"
            options={fieldOptions}
            selectedKey={actionFieldInternalName || ''}
            onChange={(_, o) => {
              const k = o ? String(o.key) : '';
              onActionFieldInternalNameChange(k);
              if (!k && captureEnabled) onCaptureEnabledChange(false);
            }}
            styles={{ root: { maxWidth: 480 } }}
            disabled={logFieldsLoading || !!logFieldsErr}
          />
        </>
      ) : null}
      <Toggle
        label="Habilitar captação de logs"
        checked={captureEnabled}
        onChange={(_, c) => onCaptureEnabledChange(!!c)}
        onText="Ativa"
        offText="Inativa"
        disabled={!canEnableCapture}
      />
      {!canEnableCapture && (
        <Text variant="small" styles={{ root: { color: '#a19f9d', fontStyle: 'italic' } }}>
          Defina a lista e o campo multilinhas para desbloquear a captação.
        </Text>
      )}
      {!customButtons.length ? (
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Configure botões no separador «Botões» para definir descrições por ação aqui.
        </Text>
      ) : (
        <Stack
          tokens={{ childrenGap: 20 }}
          styles={{
            root: captureEnabled ? undefined : { opacity: 0.55, pointerEvents: 'none' as const },
          }}
        >
          {customButtons.map((btn) => (
            <Stack
              key={btn.id}
              tokens={{ childrenGap: 8 }}
              styles={{
                root: {
                  padding: 12,
                  border: '1px solid #edebe9',
                  borderRadius: 4,
                  background: '#faf9f8',
                },
              }}
            >
              <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
                {btn.label || btn.id}{' '}
                <span style={{ color: '#605e5c', fontWeight: 400 }}>({btn.id})</span>
              </Text>
              <ListPageRichQuillEditor
                value={descriptionsHtmlByButtonId[btn.id] ?? ''}
                onChange={(html) => onDescriptionChange(btn.id, html)}
                placeholder="Descreva o que esta ação representa no registo de log…"
                permissions={LOG_QUILL_PERMISSIONS}
              />
            </Stack>
          ))}
        </Stack>
      )}
    </Stack>
  );
}
