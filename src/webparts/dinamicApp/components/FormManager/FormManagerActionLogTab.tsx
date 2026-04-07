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
  ChoiceGroup,
  IChoiceGroupOption,
  TextField,
  Separator,
  Checkbox,
  DefaultButton,
} from '@fluentui/react';
import { ListsService, FieldsService } from '../../../../services';
import type { IListSummary, IFieldMetadata, IGroupDetails } from '../../../../services';
import type {
  IFormCustomButtonConfig,
  TFormHistoryPresentationKind,
  TFormHistoryButtonKind,
} from '../../core/config/types/formManager';
import { FORM_BUILTIN_HISTORY_BUTTON_ID } from '../../core/config/types/formManager';
import { ListPageRichQuillEditor } from '../ListPage/ListPageRichQuillEditor';

const LOG_QUILL_PERMISSIONS = {
  allowHeaders: true,
  allowLists: true,
  allowLinks: true,
  allowImages: false,
  allowVideoEmbed: false,
};

const HISTORY_BUTTON_KIND_OPTIONS: IChoiceGroupOption[] = [
  { key: 'text', text: 'Só texto' },
  { key: 'icon', text: 'Só ícone' },
  { key: 'iconAndText', text: 'Ícone e texto' },
];

function normSpGroupTitle(s: string): string {
  return s.trim().toLowerCase();
}

function normListGuid(g: string | undefined): string {
  if (!g) return '';
  return g.replace(/[{}]/g, '').toLowerCase();
}

export interface IFormManagerActionLogTabProps {
  historyEnabled: boolean;
  onHistoryEnabledChange: (enabled: boolean) => void;
  historyPresentationKind: TFormHistoryPresentationKind;
  onHistoryPresentationKindChange: (v: TFormHistoryPresentationKind) => void;
  historyButtonKind: TFormHistoryButtonKind;
  onHistoryButtonKindChange: (v: TFormHistoryButtonKind) => void;
  historyButtonLabel: string;
  onHistoryButtonLabelChange: (v: string) => void;
  historyButtonIcon: string;
  onHistoryButtonIconChange: (v: string) => void;
  historyPanelSubtitle: string;
  onHistoryPanelSubtitleChange: (v: string) => void;
  historyGroupTitles: string[];
  onHistoryGroupTitlesChange: (titles: string[]) => void;
  siteGroups: IGroupDetails[];
  siteGroupsSorted: IGroupDetails[];
  siteGroupsLoading: boolean;
  siteGroupsErr?: string;
  onRetryLoadSiteGroups: () => void;
  captureEnabled: boolean;
  onCaptureEnabledChange: (enabled: boolean) => void;
  listTitle: string;
  onListTitleChange: (title: string) => void;
  actionFieldInternalName: string;
  onActionFieldInternalNameChange: (internalName: string) => void;
  descriptionsHtmlByButtonId: Record<string, string>;
  onDescriptionChange: (buttonId: string, html: string) => void;
  customButtons: IFormCustomButtonConfig[];
  /** Título da lista principal do formulário (origem dos dados). */
  primaryListTitle: string;
  sourceListLookupFieldInternalName: string;
  onSourceListLookupFieldInternalNameChange: (internalName: string) => void;
}

export function FormManagerActionLogTabContent(props: IFormManagerActionLogTabProps): JSX.Element {
  const {
    historyEnabled,
    onHistoryEnabledChange,
    historyPresentationKind,
    onHistoryPresentationKindChange,
    historyButtonKind,
    onHistoryButtonKindChange,
    historyButtonLabel,
    onHistoryButtonLabelChange,
    historyButtonIcon,
    onHistoryButtonIconChange,
    historyPanelSubtitle,
    onHistoryPanelSubtitleChange,
    historyGroupTitles,
    onHistoryGroupTitlesChange,
    siteGroups,
    siteGroupsSorted,
    siteGroupsLoading,
    siteGroupsErr,
    onRetryLoadSiteGroups,
    captureEnabled,
    onCaptureEnabledChange,
    listTitle,
    onListTitleChange,
    actionFieldInternalName,
    onActionFieldInternalNameChange,
    descriptionsHtmlByButtonId,
    onDescriptionChange,
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

  const showIconFields = historyButtonKind === 'icon' || historyButtonKind === 'iconAndText';
  const showLabelField = historyButtonKind === 'text' || historyButtonKind === 'iconAndText';
  const showAriaOnlyLabel = historyButtonKind === 'icon';

  const logDescBlocks = captureEnabled && (customButtons.length > 0 || historyEnabled);

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { marginTop: 12 } }}>
      <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
        Histórico de versões (SharePoint)
      </Text>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Botão integrado no formulário para ver o histórico de versões do item. Não é necessário criar um botão
        personalizado com tipo «Histórico» na aba Botões.
      </Text>
      <Toggle
        label="Ativar histórico de versões"
        checked={historyEnabled}
        onChange={(_, c) => onHistoryEnabledChange(!!c)}
        onText="Ativo"
        offText="Inativo"
      />
      {historyEnabled && (
        <Stack tokens={{ childrenGap: 12 }} styles={{ root: { maxWidth: 520 } }}>
          <Dropdown
            label="Abrir histórico como"
            options={[
              { key: 'panel', text: 'Painel lateral' },
              { key: 'modal', text: 'Modal' },
              { key: 'collapse', text: 'Secção no formulário (abaixo dos botões)' },
            ]}
            selectedKey={historyPresentationKind}
            onChange={(_, o) =>
              o && onHistoryPresentationKindChange(String(o.key) as TFormHistoryPresentationKind)
            }
          />
          <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Aparência do botão no formulário</Text>
          <ChoiceGroup
            options={HISTORY_BUTTON_KIND_OPTIONS}
            selectedKey={historyButtonKind}
            onChange={(_, o) =>
              o && onHistoryButtonKindChange(String(o.key) as TFormHistoryButtonKind)
            }
          />
          {showLabelField && (
            <TextField
              label="Texto do botão"
              value={historyButtonLabel}
              onChange={(_, v) => onHistoryButtonLabelChange(v ?? '')}
              placeholder="Histórico"
            />
          )}
          {showAriaOnlyLabel && (
            <TextField
              label="Nome acessível (tooltip / leitor de ecrã)"
              value={historyButtonLabel}
              onChange={(_, v) => onHistoryButtonLabelChange(v ?? '')}
              placeholder="Histórico"
            />
          )}
          {showIconFields && (
            <TextField
              label="Ícone Fluent (nome)"
              description="Ex.: History, Clock, TimelineProgress. Em «só ícone», use o subtítulo abaixo como tooltip."
              value={historyButtonIcon}
              onChange={(_, v) => onHistoryButtonIconChange(v ?? '')}
              placeholder="History"
            />
          )}
          <TextField
            label="Subtítulo / ajuda (painel de histórico e tooltip)"
            multiline
            rows={2}
            value={historyPanelSubtitle}
            onChange={(_, v) => onHistoryPanelSubtitleChange(v ?? '')}
          />
          <Text variant="small" styles={{ root: { fontWeight: 600, marginTop: 8 } }}>
            Grupos do SharePoint
          </Text>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Só utilizadores que pertençam a pelo menos um dos grupos marcados vêem o botão de histórico. Vazio = todos.
            Os títulos coincidem com os grupos devolvidos ao formulário em tempo de execução.
          </Text>
          {siteGroupsLoading && <Spinner label="A carregar grupos do site…" />}
          {siteGroupsErr && (
            <>
              <MessageBar messageBarType={MessageBarType.warning}>{siteGroupsErr}</MessageBar>
              <DefaultButton text="Tentar carregar grupos novamente" onClick={onRetryLoadSiteGroups} />
            </>
          )}
          {!siteGroupsLoading ? (
            <Stack
              tokens={{ childrenGap: 6 }}
              styles={{
                root: {
                  maxHeight: 240,
                  overflowY: 'auto',
                  border: '1px solid #edebe9',
                  borderRadius: 4,
                  padding: 8,
                },
              }}
            >
              {(historyGroupTitles ?? [])
                .filter(
                  (t) =>
                    !siteGroups.some((g) => normSpGroupTitle(g.Title) === normSpGroupTitle(t))
                )
                .map((t, oi) => (
                  <Checkbox
                    key={`hist-orphan-grp-${oi}-${t}`}
                    label={`${t} (guardado; não na lista do site)`}
                    checked
                    onChange={(_, c) => {
                      if (c) return;
                      const cur = historyGroupTitles ?? [];
                      const n = normSpGroupTitle(t);
                      const next = cur.filter((x) => normSpGroupTitle(x) !== n);
                      onHistoryGroupTitlesChange(next);
                    }}
                  />
                ))}
              {siteGroupsSorted.map((g) => {
                const cur = historyGroupTitles ?? [];
                const n = normSpGroupTitle(g.Title);
                const checked = cur.some((x) => normSpGroupTitle(x) === n);
                return (
                  <Checkbox
                    key={g.Id}
                    label={g.Title}
                    title={g.Description || undefined}
                    checked={checked}
                    onChange={(_, c) => {
                      let next: string[];
                      if (c) {
                        next = checked ? cur : cur.concat([g.Title]);
                      } else {
                        next = cur.filter((x) => normSpGroupTitle(x) !== n);
                      }
                      onHistoryGroupTitlesChange(next);
                    }}
                  />
                );
              })}
              {!siteGroupsSorted.length && !(historyGroupTitles ?? []).length && (
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  Nenhum grupo no site.
                </Text>
              )}
            </Stack>
          ) : null}
        </Stack>
      )}

      <Separator styles={{ root: { marginTop: 8 } }} />

      <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
        Lista de auditoria (registo por clique)
      </Text>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Escolha a lista de registo, um campo multilinhas para o texto da ação e um lookup que aponte para a mesma
        lista principal do formulário (vínculo ao item). Só então pode ativar a captação.
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
          {primaryListLoading && (
            <Spinner label="A resolver a lista principal do formulário…" />
          )}
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
              logFieldsLoading ||
              !!logFieldsErr ||
              primaryListLoading ||
              !primaryListId
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
      {!logDescBlocks ? (
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          {captureEnabled
            ? 'Ative o histórico acima ou configure botões no separador «Botões» para definir textos de registo aqui.'
            : 'Com a captação ativa, pode definir o texto de registo por botão (incluindo o botão de histórico integrado).'}
        </Text>
      ) : (
        <Stack
          tokens={{ childrenGap: 20 }}
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
