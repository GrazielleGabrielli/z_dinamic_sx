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
  Separator,
} from '@fluentui/react';
import { ListsService, FieldsService } from '../../../../services';
import type { IListSummary, IFieldMetadata } from '../../../../services';
import type { IFormCustomButtonConfig, TFormCustomButtonPaletteSlot } from '../../core/config/types/formManager';
import { FORM_BUILTIN_HISTORY_BUTTON_ID } from '../../core/config/types/formManager';
import { ListPageRichQuillEditor } from '../ListPage/ListPageRichQuillEditor';
import { FormManagerCollapseSection } from './FormManagerComponentsTab';
import { ThemePaletteSlotDropdown } from './ThemePaletteSlotDropdown';

const LOG_QUILL_PERMISSIONS = {
  allowHeaders: true,
  allowLists: true,
  allowLinks: true,
  allowImages: false,
  allowVideoEmbed: false,
};

function normListGuid(g: string | undefined): string {
  if (!g) return '';
  return g.replace(/[{}]/g, '').toLowerCase();
}

const LOG_SECTION_IDS = {
  list: 'logList',
  texts: 'logTexts',
} as const;

export interface IFormManagerActionLogTabProps {
  historyEnabled: boolean;
  captureEnabled: boolean;
  onCaptureEnabledChange: (enabled: boolean) => void;
  listTitle: string;
  onListTitleChange: (title: string) => void;
  actionFieldInternalName: string;
  onActionFieldInternalNameChange: (internalName: string) => void;
  descriptionsHtmlByButtonId: Record<string, string>;
  onDescriptionChange: (buttonId: string, html: string) => void;
  descriptionPaletteSlotByButtonId: Record<string, TFormCustomButtonPaletteSlot>;
  onDescriptionPaletteSlotChange: (buttonId: string, slot: TFormCustomButtonPaletteSlot) => void;
  customButtons: IFormCustomButtonConfig[];
  /** Título da lista principal do formulário (origem dos dados). */
  primaryListTitle: string;
  sourceListLookupFieldInternalName: string;
  onSourceListLookupFieldInternalNameChange: (internalName: string) => void;
}

export function FormManagerActionLogTabContent(props: IFormManagerActionLogTabProps): JSX.Element {
  const {
    historyEnabled,
    captureEnabled,
    onCaptureEnabledChange,
    listTitle,
    onListTitleChange,
    actionFieldInternalName,
    onActionFieldInternalNameChange,
    descriptionsHtmlByButtonId,
    onDescriptionChange,
    descriptionPaletteSlotByButtonId,
    onDescriptionPaletteSlotChange,
    customButtons,
    primaryListTitle,
    sourceListLookupFieldInternalName,
    onSourceListLookupFieldInternalNameChange,
  } = props;
  const [lists, setLists] = useState<IListSummary[]>([]);
  const [listsLoading, setListsLoading] = useState(false);
  const [listsErr, setListsErr] = useState<string | undefined>(undefined);
  const [logListFields, setLogListFields] = useState<IFieldMetadata[]>([]);
  const [logFieldsLoading, setLogFieldsLoading] = useState(false);
  const [logFieldsErr, setLogFieldsErr] = useState<string | undefined>(undefined);
  const [primaryListId, setPrimaryListId] = useState<string | undefined>(undefined);
  const [primaryListLoading, setPrimaryListLoading] = useState(false);
  const [openSections, setOpenSections] = useState<Record<string, boolean>>({});

  const toggleSection = (id: string): void => {
    setOpenSections((prev) => ({ ...prev, [id]: !prev[id] }));
  };
  const isSectionOpen = (id: string): boolean => openSections[id] === true;

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

  useEffect(() => {
    const t = primaryListTitle.trim();
    if (!t) {
      setPrimaryListId(undefined);
      return;
    }
    setPrimaryListLoading(true);
    const ls = new ListsService();
    ls.getListByTitle(t)
      .then((m) => {
        setPrimaryListId(m.Id);
        setPrimaryListLoading(false);
      })
      .catch(() => {
        setPrimaryListId(undefined);
        setPrimaryListLoading(false);
      });
  }, [primaryListTitle]);

  const multilineFields = useMemo(
    () => logListFields.filter((f) => f.MappedType === 'multiline' && !f.Hidden && !f.ReadOnlyField),
    [logListFields]
  );

  const linkLookupFields = useMemo(() => {
    if (!primaryListId) return [];
    const target = normListGuid(primaryListId);
    return logListFields.filter(
      (f) =>
        f.MappedType === 'lookup' &&
        !f.Hidden &&
        !f.ReadOnlyField &&
        !f.AllowMultipleValues &&
        normListGuid(f.LookupList) === target
    );
  }, [logListFields, primaryListId]);

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

  const linkFieldOptions: IDropdownOption[] = useMemo(() => {
    const base: IDropdownOption[] = [{ key: '', text: '—' }];
    const known = new Set(linkLookupFields.map((f) => f.InternalName));
    if (sourceListLookupFieldInternalName && !known.has(sourceListLookupFieldInternalName)) {
      base.push({
        key: sourceListLookupFieldInternalName,
        text: `${sourceListLookupFieldInternalName} (referência guardada)`,
      });
    }
    return base.concat(
      linkLookupFields.map((f) => ({
        key: f.InternalName,
        text: `${f.Title} (${f.InternalName})`,
      }))
    );
  }, [linkLookupFields, sourceListLookupFieldInternalName]);

  const listOptions: IDropdownOption[] = useMemo(
    () => [{ key: '', text: '— nenhuma —' }, ...lists.map((l) => ({ key: l.Title, text: l.Title }))],
    [lists]
  );

  const canEnableCapture = Boolean(
    listTitle.trim() &&
      actionFieldInternalName.trim() &&
      sourceListLookupFieldInternalName.trim()
  );

  const logDescBlocks = captureEnabled && (customButtons.length > 0 || historyEnabled);

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 12 } }}>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Expanda cada secção para configurar. O botão de histórico e a forma de abrir (painel, modal, secção)
        configuram-se na aba «Componentes». Por defeito as secções aqui vêm fechadas.
      </Text>

      <FormManagerCollapseSection
        title="Lista de registo e captação"
        isOpen={isSectionOpen(LOG_SECTION_IDS.list)}
        onToggle={() => toggleSection(LOG_SECTION_IDS.list)}
      >
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Lista SharePoint, campo multilinhas, lookup à lista principal e ativar gravação de logs.
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
              }}
              styles={{ root: { maxWidth: 480 } }}
              disabled={logFieldsLoading || !!logFieldsErr}
            />
            {primaryListLoading && <Spinner label="A resolver a lista principal do formulário…" />}
            {!primaryListTitle.trim() && (
              <MessageBar messageBarType={MessageBarType.info}>
                Indique o título da lista principal no separador «Geral» (origem dos dados) para escolher o lookup de
                vínculo ao item.
              </MessageBar>
            )}
            {!!primaryListTitle.trim() && !primaryListLoading && !primaryListId && (
              <MessageBar messageBarType={MessageBarType.warning}>
                Não foi possível obter a lista principal «{primaryListTitle}». Confira o título no separador «Geral».
              </MessageBar>
            )}
            <Dropdown
              label="Lookup para a lista principal (vínculo ao item)"
              options={linkFieldOptions}
              selectedKey={sourceListLookupFieldInternalName || ''}
              onChange={(_, o) => {
                const k = o ? String(o.key) : '';
                onSourceListLookupFieldInternalNameChange(k);
              }}
              styles={{ root: { maxWidth: 480 } }}
              disabled={
                logFieldsLoading || !!logFieldsErr || primaryListLoading || !primaryListId
              }
            />
            {!logFieldsLoading &&
              !logFieldsErr &&
              primaryListId &&
              linkLookupFields.length === 0 && (
                <MessageBar messageBarType={MessageBarType.warning}>
                  Não há coluna de lookup nesta lista de registo que aponte para «{primaryListTitle}». Crie uma coluna
                  lookup para essa lista.
                </MessageBar>
              )}
          </>
        ) : null}
        <Separator />
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
            Defina a lista, o campo multilinhas e o lookup de vínculo à lista principal para desbloquear a captação.
          </Text>
        )}
      </FormManagerCollapseSection>

      <FormManagerCollapseSection
        title="Textos de registo por botão"
        isOpen={isSectionOpen(LOG_SECTION_IDS.texts)}
        onToggle={() => toggleSection(LOG_SECTION_IDS.texts)}
      >
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Editor HTML por botão (histórico integrado e botões personalizados) quando a captação está ativa.
        </Text>
        {!logDescBlocks ? (
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            {captureEnabled
              ? 'Ative o histórico na aba «Componentes» ou configure botões no separador «Botões» para editar textos aqui.'
              : 'Ative a captação na secção anterior e tenha histórico ou botões configurados para editar os textos de registo.'}
          </Text>
        ) : (
          <Stack
            tokens={{ childrenGap: 16 }}
            styles={{
              root: captureEnabled ? undefined : { opacity: 0.55, pointerEvents: 'none' as const },
            }}
          >
            {historyEnabled && (
              <Stack
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
                  Botão de histórico (integrado){' '}
                  <span style={{ color: '#605e5c', fontWeight: 400 }}>({FORM_BUILTIN_HISTORY_BUTTON_ID})</span>
                </Text>
                <ThemePaletteSlotDropdown
                  label="Cor do registo (tema)"
                  selectedKey={
                    descriptionPaletteSlotByButtonId[FORM_BUILTIN_HISTORY_BUTTON_ID] ?? 'themePrimary'
                  }
                  onChange={(slot) => onDescriptionPaletteSlotChange(FORM_BUILTIN_HISTORY_BUTTON_ID, slot)}
                />
                <ListPageRichQuillEditor
                  value={descriptionsHtmlByButtonId[FORM_BUILTIN_HISTORY_BUTTON_ID] ?? ''}
                  onChange={(html) => onDescriptionChange(FORM_BUILTIN_HISTORY_BUTTON_ID, html)}
                  placeholder="Texto gravado no registo de log ao abrir o histórico…"
                  permissions={LOG_QUILL_PERMISSIONS}
                />
              </Stack>
            )}
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
                <ThemePaletteSlotDropdown
                  label="Cor do registo (tema)"
                  selectedKey={descriptionPaletteSlotByButtonId[btn.id] ?? 'themePrimary'}
                  onChange={(slot) => onDescriptionPaletteSlotChange(btn.id, slot)}
                />
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
      </FormManagerCollapseSection>
    </Stack>
  );
}
