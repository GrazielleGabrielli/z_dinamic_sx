import * as React from 'react';
import { useState, useEffect, useMemo, useCallback } from 'react';
import {
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  Spinner,
  MessageBar,
  MessageBarType,
  IconButton,
  Icon,
} from '@fluentui/react';
import { FieldsService, ListsService } from '../../../../services';
import type { IFieldMetadata, IListSummary } from '../../../../services';
import type {
  IFormFieldConfig,
  IFormLinkedChildFormConfig,
  IFormManagerAttachmentLibraryConfig,
  TFormAttachmentStorageKind,
  TLinkedChildAttachmentStorageKind,
  TLinkedChildRowsPresentationKind,
} from '../../core/config/types/formManager';
import {
  FORM_ATTACHMENTS_FIELD_INTERNAL,
  FORM_FIXOS_STEP_ID,
  FORM_OCULTOS_STEP_ID,
} from '../../core/config/types/formManager';
import { newLinkedChildFormConfig } from '../../core/config/utils';
import {
  addRootSibling,
  loadFolderTreeFromAttachmentLibrary,
} from '../../core/formManager/attachmentFolderTree';
import { FormManagerFolderTreeEditor } from './FormManagerFolderTreeEditor';
import { FormFieldRulesPanel } from './FormFieldRulesPanel';
import { FormManagerLinkedChildConditionalRulesBlock } from './FormManagerLinkedChildConditionalRulesBlock';
import { FormManagerLinkedChildPresentationPreview } from './FormManagerLinkedChildPresentationPreview';
import { buildFieldUiRules, mergeFieldRules } from '../../core/formManager/formManagerVisualModel';

const MAX_LINKED = 10;

const ROWS_PRESENTATION_OPTIONS: { key: TLinkedChildRowsPresentationKind; text: string }[] = [
  { key: 'stack', text: 'Blocos (em coluna)' },
  { key: 'table', text: 'Tabela' },
  { key: 'compact', text: 'Compacto' },
  { key: 'cards', text: 'Cartões' },
];

const LINKED_CHILD_STORAGE_OPTIONS: { key: TLinkedChildAttachmentStorageKind; text: string }[] = [
  { key: 'none', text: 'Sem anexos neste bloco' },
  { key: 'itemAttachments', text: 'Anexos nativos (lista filha)' },
  {
    key: 'documentLibraryInheritMain',
    text: 'Biblioteca da aba Anexos (pastas herdadas; Lookup à lista filha)',
  },
  { key: 'documentLibraryCustom', text: 'Outra biblioteca (estrutura própria)' },
];

const EXCLUDE_INTERNALS = new Set(['Attachments', 'ContentType', 'ContentTypeId']);

const ORDER_DROPDOWN: IDropdownOption[] = [
  { key: '__auto', text: 'Automático (0)' },
  ...Array.from({ length: 31 }, (_, i) => ({
    key: String(i),
    text: String(i),
  })),
];

const MIN_ROWS_DROPDOWN: IDropdownOption[] = Array.from({ length: 21 }, (_, i) => ({
  key: String(i),
  text: i === 0 ? '0 (sem mínimo obrigatório)' : String(i),
}));

const MAX_ROWS_DROPDOWN: IDropdownOption[] = [
  { key: '__none', text: 'Sem limite' },
  ...Array.from({ length: 50 }, (_, i) => ({
    key: String(i + 1),
    text: String(i + 1),
  })),
];

function normListGuid(g: string | undefined): string {
  if (!g) return '';
  return g.replace(/[{}]/g, '').toLowerCase();
}

/** Guid da lista referenciada por LookupList (REST pode devolver `{guid}`, URL ou só o guid). */
function listGuidFromLookupListField(raw: string | undefined): string {
  if (!raw) return '';
  const t = raw.trim();
  const inBraces = t.match(/\{([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\}/i);
  if (inBraces) return normListGuid(inBraces[1]);
  const plain = t.match(
    /([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i
  );
  if (plain) return normListGuid(plain[1]);
  return normListGuid(t);
}

function buildParentLookupDropdownOptions(
  meta: IFieldMetadata[],
  primaryListIdNorm: string,
  savedParentLookupInternalName: string
): IDropdownOption[] {
  const lookupFields = meta.filter((m) => m.MappedType === 'lookup' && m.LookupList);
  const filtered = primaryListIdNorm
    ? lookupFields.filter(
        (m) => listGuidFromLookupListField(m.LookupList) === primaryListIdNorm
      )
    : lookupFields;
  const opts = filtered.map((m) => ({ key: m.InternalName, text: `${m.Title} (${m.InternalName})` }));
  const saved = savedParentLookupInternalName.trim();
  if (!saved) return opts;
  const keys = new Set(opts.map((o) => String(o.key)));
  if (keys.has(saved)) return opts;
  const baseSaved = saved.replace(/Id$/i, '');
  const fromMeta = lookupFields.find(
    (m) =>
      m.InternalName === saved ||
      m.InternalName === `${saved}Id` ||
      m.InternalName.replace(/Id$/i, '') === baseSaved
  );
  if (fromMeta && !keys.has(fromMeta.InternalName)) {
    return opts.concat([
      { key: fromMeta.InternalName, text: `${fromMeta.Title} (${fromMeta.InternalName})` },
    ]);
  }
  return opts.concat([
    {
      key: saved,
      text: `${saved} (guardado na configuração)`,
    },
  ]);
}

function parentLookupSelectedKeyForDropdown(
  meta: IFieldMetadata[],
  savedInternalName: string
): string {
  const t = savedInternalName.trim();
  if (!t) return '';
  const lookups = meta.filter((m) => m.MappedType === 'lookup' && m.LookupList);
  if (lookups.some((m) => m.InternalName === t)) return t;
  const base = t.replace(/Id$/i, '');
  const hit = lookups.find((m) => m.InternalName.replace(/Id$/i, '') === base);
  return hit?.InternalName ?? t;
}

function isSelectableChildField(m: IFieldMetadata): boolean {
  return !EXCLUDE_INTERNALS.has(m.InternalName) && m.InternalName !== FORM_ATTACHMENTS_FIELD_INTERNAL;
}

function newId(): string {
  return `lcf_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
}

function patchChildById(
  arr: IFormLinkedChildFormConfig[],
  id: string,
  patch: Partial<IFormLinkedChildFormConfig>
): IFormLinkedChildFormConfig[] {
  return arr.map((c) => (c.id === id ? { ...c, ...patch } : c));
}

function setMainStepFieldNames(
  cfg: IFormLinkedChildFormConfig,
  fieldNames: string[]
): IFormLinkedChildFormConfig {
  const steps = (cfg.steps ?? []).map((s) =>
    s.id === 'main' ? { ...s, fieldNames: fieldNames.slice() } : { ...s, fieldNames: s.fieldNames.slice() }
  );
  return { ...cfg, steps };
}

function ensureFieldConfigsForNames(
  cfg: IFormLinkedChildFormConfig,
  names: string[]
): IFormLinkedChildFormConfig {
  const existing = new Map(cfg.fields.map((f) => [f.internalName, f]));
  const fields: IFormFieldConfig[] = [];
  for (let i = 0; i < names.length; i++) {
    const n = names[i];
    fields.push(existing.get(n) ?? { internalName: n });
  }
  return { ...cfg, fields };
}

function listDropdownOptionsWithLegacy(
  base: IDropdownOption[],
  currentTitle: string
): IDropdownOption[] {
  const t = currentTitle.trim();
  if (!t) return base;
  const has = base.some((o) => String(o.key) === t);
  if (has) return base;
  return base.concat([{ key: t, text: `${t} (não encontrada — reverifique)` }]);
}

export interface IFormManagerLinkedChildFormsTabProps {
  primaryListTitle: string;
  linkedChildForms: IFormLinkedChildFormConfig[];
  onLinkedChildFormsChange: React.Dispatch<React.SetStateAction<IFormLinkedChildFormConfig[]>>;
  mainAttachmentStorageKind: TFormAttachmentStorageKind | undefined;
  mainAttachmentLibraryFromPanel: IFormManagerAttachmentLibraryConfig | undefined;
}

function childFolderStepOptions(cfg: IFormLinkedChildFormConfig): { id: string; title: string }[] {
  return (cfg.steps ?? [])
    .filter((s) => s.id !== FORM_OCULTOS_STEP_ID && s.id !== FORM_FIXOS_STEP_ID)
    .map((s) => ({ id: s.id, title: s.title }));
}

interface ILinkedChildCardPanels {
  card: boolean;
  connection: boolean;
  presentation: boolean;
  attachments: boolean;
  fields: boolean;
  rules: boolean;
}

const DEFAULT_LINKED_CHILD_CARD_PANELS: ILinkedChildCardPanels = {
  card: false,
  connection: false,
  presentation: false,
  attachments: false,
  fields: false,
  rules: false,
};

function LinkedChildTabCollapseSection(props: {
  title: string;
  expanded: boolean;
  onToggle: () => void;
  indent?: boolean;
  children: React.ReactNode;
}): JSX.Element {
  const { title, expanded, onToggle, indent, children } = props;
  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <Stack
        horizontal
        verticalAlign="center"
        tokens={{ childrenGap: 8 }}
        styles={{
          root: {
            cursor: 'pointer',
            userSelect: 'none',
          },
        }}
        onClick={onToggle}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            onToggle();
          }
        }}
        role="button"
        tabIndex={0}
        aria-expanded={expanded}
      >
        <Icon
          iconName={expanded ? 'ChevronDown' : 'ChevronRight'}
          styles={{ root: { fontSize: 14, color: '#323130', flexShrink: 0 } }}
        />
        <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
          {title}
        </Text>
      </Stack>
      {expanded && (
        <Stack
          tokens={{ childrenGap: 10 }}
          styles={
            indent
              ? { root: { paddingLeft: 22, borderLeft: '2px solid #edebe9' } }
              : undefined
          }
        >
          {children}
        </Stack>
      )}
    </Stack>
  );
}

export function FormManagerLinkedChildFormsTabContent(props: IFormManagerLinkedChildFormsTabProps): JSX.Element {
  const {
    primaryListTitle,
    linkedChildForms,
    onLinkedChildFormsChange,
    mainAttachmentStorageKind,
    mainAttachmentLibraryFromPanel,
  } = props;
  const fieldsService = useMemo(() => new FieldsService(), []);
  const listsService = useMemo(() => new ListsService(), []);
  const [siteLists, setSiteLists] = useState<IListSummary[]>([]);
  const [listsLoading, setListsLoading] = useState(false);
  const [childMetaById, setChildMetaById] = useState<Record<string, IFieldMetadata[]>>({});
  const [childMetaLoading, setChildMetaLoading] = useState<Record<string, boolean>>({});
  const [inheritLibFieldOpts, setInheritLibFieldOpts] = useState<Record<string, IDropdownOption[]>>({});
  const [inheritLibLoading, setInheritLibLoading] = useState<Record<string, boolean>>({});
  const [customLibFieldOpts, setCustomLibFieldOpts] = useState<Record<string, IDropdownOption[]>>({});
  const [customLibLoading, setCustomLibLoading] = useState<Record<string, boolean>>({});
  const [linkedChildCardPanels, setLinkedChildCardPanels] = useState<
    Record<string, ILinkedChildCardPanels>
  >({});

  useEffect(() => {
    setListsLoading(true);
    listsService
      .getLists(false)
      .then((lists) => setSiteLists(lists))
      .catch(() => setSiteLists([]))
      .finally(() => setListsLoading(false));
  }, [listsService]);

  const childListsForDropdown = useMemo(() => {
    const lists = siteLists
      .filter((l) => !l.IsLibrary)
      .slice()
      .sort((a, b) => a.Title.localeCompare(b.Title, 'pt', { sensitivity: 'base' }));
    return lists;
  }, [siteLists]);

  const childListDropdownOptions = useMemo((): IDropdownOption[] => {
    const opts: IDropdownOption[] = [{ key: '', text: '— Selecione uma lista —' }];
    for (let i = 0; i < childListsForDropdown.length; i++) {
      const l = childListsForDropdown[i];
      opts.push({
        key: l.Title,
        text: `${l.Title} (${l.ItemCount} itens)`,
      });
    }
    return opts;
  }, [childListsForDropdown]);

  const siteLibrariesSorted = useMemo(() => {
    return siteLists
      .filter((l) => l.IsLibrary)
      .slice()
      .sort((a, b) => a.Title.localeCompare(b.Title, 'pt', { sensitivity: 'base' }));
  }, [siteLists]);

  const attachmentLibraryDropdownOptions = useMemo((): IDropdownOption[] => {
    const opts: IDropdownOption[] = [{ key: '', text: '— Selecione uma biblioteca —' }];
    for (let i = 0; i < siteLibrariesSorted.length; i++) {
      const l = siteLibrariesSorted[i];
      opts.push({
        key: l.Title,
        text: `${l.Title}${typeof l.ItemCount === 'number' ? ` (${l.ItemCount})` : ''}`,
      });
    }
    return opts;
  }, [siteLibrariesSorted]);

  const primaryListId = useMemo(() => {
    const t = primaryListTitle.trim();
    if (!t || !siteLists.length) return '';
    const hit = siteLists.find((l) => l.Title.trim().toLowerCase() === t.toLowerCase());
    return hit?.Id ? normListGuid(hit.Id) : '';
  }, [primaryListTitle, siteLists]);

  const loadChildMeta = useCallback(
    async (configId: string, listTitle: string): Promise<void> => {
      const lt = listTitle.trim();
      if (!lt) {
        setChildMetaById((prev) => {
          const next = { ...prev };
          delete next[configId];
          return next;
        });
        return;
      }
      setChildMetaLoading((prev) => ({ ...prev, [configId]: true }));
      try {
        const f = await fieldsService.getVisibleFields(lt);
        setChildMetaById((prev) => ({ ...prev, [configId]: f }));
      } catch {
        setChildMetaById((prev) => {
          const next = { ...prev };
          delete next[configId];
          return next;
        });
      } finally {
        setChildMetaLoading((prev) => ({ ...prev, [configId]: false }));
      }
    },
    [fieldsService]
  );

  useEffect(() => {
    for (let i = 0; i < linkedChildForms.length; i++) {
      const cfg = linkedChildForms[i];
      const t = cfg.listTitle.trim();
      if (!t) continue;
      void loadChildMeta(cfg.id, t);
    }
  }, [linkedChildForms, loadChildMeta]);

  useEffect(() => {
    let cancel = false;
    void (async () => {
      for (let i = 0; i < linkedChildForms.length; i++) {
        const cfg = linkedChildForms[i];
        if (cfg.childAttachmentStorageKind !== 'documentLibraryInheritMain') continue;
        const mainLibTitle = (mainAttachmentLibraryFromPanel?.libraryTitle ?? '').trim();
        const childTitle = cfg.listTitle.trim();
        if (!mainLibTitle || !childTitle) {
          if (!cancel) {
            setInheritLibFieldOpts((o) => ({ ...o, [cfg.id]: [{ key: '', text: '—' }] }));
          }
          continue;
        }
        const childList = siteLists.find((l) => l.Title.trim().toLowerCase() === childTitle.toLowerCase());
        const childGuid = childList?.Id ? normListGuid(String(childList.Id)) : '';
        if (!childGuid) {
          if (!cancel) {
            setInheritLibFieldOpts((o) => ({
              ...o,
              [cfg.id]: [{ key: '', text: '— lista filha não encontrada —' }],
            }));
          }
          continue;
        }
        if (!cancel) setInheritLibLoading((x) => ({ ...x, [cfg.id]: true }));
        try {
          const meta = await fieldsService.getVisibleFields(mainLibTitle);
          if (cancel) return;
          const opts = meta
            .filter((m) => m.MappedType === 'lookup' && m.LookupList && normListGuid(m.LookupList) === childGuid)
            .map((m) => ({ key: m.InternalName, text: `${m.Title} (${m.InternalName})` }));
          if (!opts.length) {
            setInheritLibFieldOpts((o) => ({
              ...o,
              [cfg.id]: [{ key: '', text: '— sem Lookup à lista filha nesta biblioteca —' }],
            }));
          } else {
            setInheritLibFieldOpts((o) => ({
              ...o,
              [cfg.id]: [{ key: '', text: '—' }, ...opts],
            }));
          }
        } catch {
          if (!cancel) {
            setInheritLibFieldOpts((o) => ({
              ...o,
              [cfg.id]: [{ key: '', text: '— erro ao carregar campos —' }],
            }));
          }
        } finally {
          if (!cancel) setInheritLibLoading((x) => ({ ...x, [cfg.id]: false }));
        }
      }
    })();
    return (): void => {
      cancel = true;
    };
  }, [linkedChildForms, mainAttachmentLibraryFromPanel, siteLists, fieldsService]);

  useEffect(() => {
    let cancel = false;
    void (async () => {
      for (let i = 0; i < linkedChildForms.length; i++) {
        const cfg = linkedChildForms[i];
        if (cfg.childAttachmentStorageKind !== 'documentLibraryCustom') continue;
        const libT = (cfg.childAttachmentLibrary?.libraryTitle ?? '').trim();
        const childTitle = cfg.listTitle.trim();
        if (!libT || !childTitle) {
          if (!cancel) setCustomLibFieldOpts((o) => ({ ...o, [cfg.id]: [{ key: '', text: '—' }] }));
          continue;
        }
        const childList = siteLists.find((l) => l.Title.trim().toLowerCase() === childTitle.toLowerCase());
        const childGuid = childList?.Id ? normListGuid(String(childList.Id)) : '';
        if (!childGuid) {
          if (!cancel) {
            setCustomLibFieldOpts((o) => ({
              ...o,
              [cfg.id]: [{ key: '', text: '— lista filha não encontrada —' }],
            }));
          }
          continue;
        }
        if (!cancel) setCustomLibLoading((x) => ({ ...x, [cfg.id]: true }));
        try {
          const meta = await fieldsService.getVisibleFields(libT);
          if (cancel) return;
          const opts = meta
            .filter((m) => m.MappedType === 'lookup' && m.LookupList && normListGuid(m.LookupList) === childGuid)
            .map((m) => ({ key: m.InternalName, text: `${m.Title} (${m.InternalName})` }));
          if (!opts.length) {
            setCustomLibFieldOpts((o) => ({
              ...o,
              [cfg.id]: [{ key: '', text: '— sem Lookup à lista filha —' }],
            }));
          } else {
            setCustomLibFieldOpts((o) => ({
              ...o,
              [cfg.id]: [{ key: '', text: '—' }, ...opts],
            }));
          }
        } catch {
          if (!cancel) {
            setCustomLibFieldOpts((o) => ({
              ...o,
              [cfg.id]: [{ key: '', text: '— erro ao carregar campos —' }],
            }));
          }
        } finally {
          if (!cancel) setCustomLibLoading((x) => ({ ...x, [cfg.id]: false }));
        }
      }
    })();
    return (): void => {
      cancel = true;
    };
  }, [linkedChildForms, siteLists, fieldsService]);

  const addLinked = (): void => {
    if (linkedChildForms.length >= MAX_LINKED) return;
    onLinkedChildFormsChange([...linkedChildForms, newLinkedChildFormConfig(newId())]);
  };

  const removeLinked = (id: string): void => {
    onLinkedChildFormsChange(linkedChildForms.filter((c) => c.id !== id));
    setChildMetaById((prev) => {
      const next = { ...prev };
      delete next[id];
      return next;
    });
    setInheritLibFieldOpts((prev) => {
      const next = { ...prev };
      delete next[id];
      return next;
    });
    setInheritLibLoading((prev) => {
      const next = { ...prev };
      delete next[id];
      return next;
    });
    setCustomLibFieldOpts((prev) => {
      const next = { ...prev };
      delete next[id];
      return next;
    });
    setCustomLibLoading((prev) => {
      const next = { ...prev };
      delete next[id];
      return next;
    });
  };

  const moveLinked = (from: number, to: number): void => {
    if (to < 0 || to >= linkedChildForms.length) return;
    const next = linkedChildForms.slice();
    const [m] = next.splice(from, 1);
    next.splice(to, 0, m);
    onLinkedChildFormsChange(next);
  };

  const [linkedFieldRuleTarget, setLinkedFieldRuleTarget] = useState<{
    cfgId: string;
    internalName: string;
  } | null>(null);

  return (
    <>
    <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Listas secundárias com Lookup para a lista principal. A ordem na etapa «Geral» define cada mini-formulário.
        Em «Estrutura», escolha em que etapa do passador cada bloco aparece.
      </Text>
      {!primaryListTitle.trim() && (
        <MessageBar messageBarType={MessageBarType.warning}>
          Defina primeiro a lista principal na configuração da vista (origem de dados).
        </MessageBar>
      )}
      {listsLoading && <Spinner label="A carregar listas do site…" />}
      {primaryListTitle.trim() && !listsLoading && !primaryListId && (
        <MessageBar messageBarType={MessageBarType.severeWarning}>
          Não foi encontrada uma lista com o título «{primaryListTitle.trim()}» no site. Os Lookups filtrados
          podem ficar incompletos.
        </MessageBar>
      )}
      <PrimaryButton text="Adicionar lista vinculada" onClick={addLinked} disabled={linkedChildForms.length >= MAX_LINKED} />
      {linkedChildForms.map((cfg, index) => {
        const meta = childMetaById[cfg.id] ?? [];
        const loading = childMetaLoading[cfg.id] === true;
        const listOpts = listDropdownOptionsWithLegacy(childListDropdownOptions, cfg.listTitle);
        const mainNames =
          cfg.steps?.find((s) => s.id === 'main')?.fieldNames?.slice() ?? [];
        const parentLookupSaved = cfg.parentLookupFieldInternalName.trim();
        const parentLookupResolvedKey = parentLookupSelectedKeyForDropdown(meta, cfg.parentLookupFieldInternalName);
        const lookupOpts = buildParentLookupDropdownOptions(
          meta,
          primaryListId,
          cfg.parentLookupFieldInternalName
        );
        const lookupDropdownOptions: IDropdownOption[] = lookupOpts.length
          ? [{ key: '', text: '— Selecione o campo Lookup —' }, ...lookupOpts]
          : [
              {
                key: '',
                text: cfg.listTitle.trim()
                  ? loading
                    ? 'A carregar campos…'
                    : '— Nenhum campo Lookup para a lista principal —'
                  : '— Escolha primeiro a lista filha —',
              },
            ];
        const availableToAdd = meta.filter(
          (m) =>
            isSelectableChildField(m) &&
            m.InternalName !== parentLookupSaved &&
            m.InternalName !== parentLookupResolvedKey &&
            mainNames.indexOf(m.InternalName) === -1
        );
        const addOptions: IDropdownOption[] = availableToAdd.map((m) => ({
          key: m.InternalName,
          text: `${m.Title} (${m.InternalName})`,
        }));

        const panels = linkedChildCardPanels[cfg.id] ?? DEFAULT_LINKED_CHILD_CARD_PANELS;
        const flip = (key: keyof ILinkedChildCardPanels): (() => void) => () =>
          setLinkedChildCardPanels((m) => {
            const cur = m[cfg.id] ?? DEFAULT_LINKED_CHILD_CARD_PANELS;
            return { ...m, [cfg.id]: { ...cur, [key]: !cur[key] } };
          });

        return (
          <Stack
            key={cfg.id}
            tokens={{ childrenGap: 10 }}
            styles={{
              root: {
                border: '1px solid #edebe9',
                borderRadius: 4,
                padding: 12,
                background: '#ffffff',
              },
            }}
          >
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Stack
                horizontal
                verticalAlign="center"
                tokens={{ childrenGap: 8 }}
                styles={{
                  root: {
                    cursor: 'pointer',
                    userSelect: 'none',
                    flex: 1,
                    minWidth: 0,
                  },
                }}
                onClick={flip('card')}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    flip('card')();
                  }
                }}
                role="button"
                tabIndex={0}
                aria-expanded={panels.card}
              >
                <Icon
                  iconName={panels.card ? 'ChevronDown' : 'ChevronRight'}
                  styles={{ root: { fontSize: 16, color: '#323130', flexShrink: 0 } }}
                />
                <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                  Lista vinculada {index + 1}: {cfg.listTitle.trim() || '(sem título)'}
                </Text>
              </Stack>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }} styles={{ root: { flexShrink: 0 } }}>
                <IconButton
                  iconProps={{ iconName: 'Up' }}
                  title="Mover bloco para cima"
                  disabled={index === 0}
                  onClick={(e) => {
                    e.stopPropagation();
                    moveLinked(index, index - 1);
                  }}
                />
                <IconButton
                  iconProps={{ iconName: 'Down' }}
                  title="Mover bloco para baixo"
                  disabled={index === linkedChildForms.length - 1}
                  onClick={(e) => {
                    e.stopPropagation();
                    moveLinked(index, index + 1);
                  }}
                />
                <DefaultButton
                  text="Remover bloco"
                  onClick={(e) => {
                    e.stopPropagation();
                    removeLinked(cfg.id);
                  }}
                />
              </Stack>
            </Stack>
            {panels.card && (
              <Stack tokens={{ childrenGap: 14 }} styles={{ root: { marginTop: 4 } }}>
                <LinkedChildTabCollapseSection
                  title="Lista e ligação ao principal"
                  expanded={panels.connection}
                  onToggle={flip('connection')}
                  indent
                >
                  <Dropdown
                    label="Lista filha (SharePoint)"
                    options={listOpts}
                    selectedKey={cfg.listTitle.trim() || ''}
                    onChange={(_, o) => {
                      if (!o) return;
                      const title = String(o.key);
                      onLinkedChildFormsChange(
                        patchChildById(linkedChildForms, cfg.id, {
                          listTitle: title,
                          parentLookupFieldInternalName: '',
                        })
                      );
                      void loadChildMeta(cfg.id, title);
                    }}
                    disabled={listsLoading}
                  />
                  {loading && <Spinner label="A carregar campos da lista…" />}
                  <Dropdown
                    label="Campo Lookup para a lista principal"
                    options={lookupDropdownOptions}
                    selectedKey={parentLookupResolvedKey || parentLookupSaved || ''}
                    onChange={(_, o) =>
                      o &&
                      onLinkedChildFormsChange(
                        patchChildById(linkedChildForms, cfg.id, {
                          parentLookupFieldInternalName: String(o.key),
                        })
                      )
                    }
                    disabled={
                      !cfg.listTitle.trim() ||
                      loading ||
                      (lookupOpts.length === 0 && !parentLookupSaved)
                    }
                  />
                  {cfg.title?.trim() && (
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} wrap>
                      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                        Rótulo JSON legado: «{cfg.title}»
                      </Text>
                      <DefaultButton
                        text="Usar só o nome da lista"
                        onClick={() =>
                          onLinkedChildFormsChange(
                            patchChildById(linkedChildForms, cfg.id, { title: undefined })
                          )
                        }
                      />
                    </Stack>
                  )}
                </LinkedChildTabCollapseSection>
                <LinkedChildTabCollapseSection
                  title="Apresentação no formulário"
                  expanded={panels.presentation}
                  onToggle={flip('presentation')}
                  indent
                >
                  <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
                    <Dropdown
                      label="Ordem de exibição"
                      options={ORDER_DROPDOWN}
                      selectedKey={cfg.order === undefined ? '__auto' : String(cfg.order)}
                      onChange={(_, o) => {
                        if (!o) return;
                        const k = String(o.key);
                        onLinkedChildFormsChange(
                          patchChildById(linkedChildForms, cfg.id, {
                            order: k === '__auto' ? undefined : parseInt(k, 10),
                          })
                        );
                      }}
                      styles={{ root: { minWidth: 200 } }}
                    />
                    <Dropdown
                      label="Mínimo de linhas"
                      options={MIN_ROWS_DROPDOWN}
                      selectedKey={String(cfg.minRows ?? 0)}
                      onChange={(_, o) =>
                        o &&
                        onLinkedChildFormsChange(
                          patchChildById(linkedChildForms, cfg.id, {
                            minRows: parseInt(String(o.key), 10),
                          })
                        )
                      }
                      styles={{ root: { minWidth: 200 } }}
                    />
                    <Dropdown
                      label="Máximo de linhas"
                      options={MAX_ROWS_DROPDOWN}
                      selectedKey={cfg.maxRows === undefined ? '__none' : String(cfg.maxRows)}
                      onChange={(_, o) => {
                        if (!o) return;
                        const k = String(o.key);
                        onLinkedChildFormsChange(
                          patchChildById(linkedChildForms, cfg.id, {
                            maxRows: k === '__none' ? undefined : parseInt(k, 10),
                          })
                        );
                      }}
                      styles={{ root: { minWidth: 200 } }}
                    />
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
                    <Dropdown
                      label="Apresentação das linhas"
                      options={ROWS_PRESENTATION_OPTIONS}
                      selectedKey={cfg.rowsPresentation ?? 'stack'}
                      onChange={(_, o) => {
                        if (!o) return;
                        const k = String(o.key) as TLinkedChildRowsPresentationKind;
                        onLinkedChildFormsChange(
                          patchChildById(linkedChildForms, cfg.id, {
                            rowsPresentation: k === 'stack' ? undefined : k,
                          })
                        );
                      }}
                      styles={{ root: { minWidth: 260 } }}
                    />
                  </Stack>
                  <FormManagerLinkedChildPresentationPreview
                    cfg={cfg}
                    fieldMeta={meta}
                    presentationKind={cfg.rowsPresentation ?? 'stack'}
                  />
                </LinkedChildTabCollapseSection>
                <LinkedChildTabCollapseSection
                  title="Anexos e biblioteca (por linha)"
                  expanded={panels.attachments}
                  onToggle={flip('attachments')}
                  indent
                >
                  {(() => {
                    const mainCanInherit =
                      mainAttachmentStorageKind === 'documentLibrary' &&
                      !!(mainAttachmentLibraryFromPanel?.libraryTitle ?? '').trim() &&
                      !!(mainAttachmentLibraryFromPanel?.sourceListLookupFieldInternalName ?? '').trim();
                    const storageOpts: IDropdownOption[] = LINKED_CHILD_STORAGE_OPTIONS.filter(
                      (o) => o.key !== 'documentLibraryInheritMain' || mainCanInherit
                    ).map((o) => ({ key: o.key, text: o.text }));
                    const sk: TLinkedChildAttachmentStorageKind = cfg.childAttachmentStorageKind ?? 'none';
                    const stepOpts = childFolderStepOptions(cfg);
                    const libBase = cfg.childAttachmentLibrary ?? {
                      libraryTitle: '',
                      sourceListLookupFieldInternalName: '',
                      folderTree: addRootSibling([]),
                    };
                    return (
                      <Stack tokens={{ childrenGap: 10 }}>
                        <Dropdown
                    label="Modo"
                    options={storageOpts}
                    selectedKey={sk}
                    onChange={(_, o) => {
                      if (!o) return;
                      const k = String(o.key) as TLinkedChildAttachmentStorageKind;
                      if (k === 'none') {
                        onLinkedChildFormsChange(
                          patchChildById(linkedChildForms, cfg.id, {
                            childAttachmentStorageKind: 'none',
                            childAttachmentLibraryLookupToChildListField: undefined,
                            childAttachmentLibrary: undefined,
                          })
                        );
                      } else if (k === 'itemAttachments') {
                        onLinkedChildFormsChange(
                          patchChildById(linkedChildForms, cfg.id, {
                            childAttachmentStorageKind: 'itemAttachments',
                            childAttachmentLibraryLookupToChildListField: undefined,
                            childAttachmentLibrary: undefined,
                          })
                        );
                      } else if (k === 'documentLibraryInheritMain') {
                        onLinkedChildFormsChange(
                          patchChildById(linkedChildForms, cfg.id, {
                            childAttachmentStorageKind: 'documentLibraryInheritMain',
                            childAttachmentLibrary: undefined,
                            childAttachmentLibraryLookupToChildListField: '',
                          })
                        );
                      } else if (k === 'documentLibraryCustom') {
                        onLinkedChildFormsChange(
                          patchChildById(linkedChildForms, cfg.id, {
                            childAttachmentStorageKind: 'documentLibraryCustom',
                            childAttachmentLibrary: {
                              libraryTitle: libBase.libraryTitle ?? '',
                              sourceListLookupFieldInternalName: '',
                              folderTree: libBase.folderTree?.length
                                ? libBase.folderTree
                                : addRootSibling([]),
                            },
                            childAttachmentLibraryLookupToChildListField: undefined,
                          })
                        );
                      }
                    }}
                  />
                  {sk === 'documentLibraryInheritMain' && !mainCanInherit && (
                    <MessageBar messageBarType={MessageBarType.warning}>
                      Para herdar a biblioteca da aba Anexos, configure lá «Biblioteca de documentos» com título da
                      biblioteca e Lookup para a lista principal.
                    </MessageBar>
                  )}
                  {sk === 'documentLibraryInheritMain' && mainCanInherit && (
                    <Stack tokens={{ childrenGap: 8 }}>
                      {inheritLibLoading[cfg.id] === true && <Spinner label="A carregar campos da biblioteca…" />}
                      <Dropdown
                        label="Campo Lookup na biblioteca de Anexos (referência à lista filha)"
                        options={
                          inheritLibFieldOpts[cfg.id] ?? [
                            {
                              key: '',
                              text:
                                inheritLibLoading[cfg.id] === true
                                  ? 'A carregar…'
                                  : '— defina lista filha e título da biblioteca em Anexos —',
                            },
                          ]
                        }
                        selectedKey={(cfg.childAttachmentLibraryLookupToChildListField ?? '').trim()}
                        onChange={(_, o) =>
                          o &&
                          onLinkedChildFormsChange(
                            patchChildById(linkedChildForms, cfg.id, {
                              childAttachmentLibraryLookupToChildListField: String(o.key),
                            })
                          )
                        }
                        disabled={inheritLibLoading[cfg.id] === true}
                      />
                    </Stack>
                  )}
                  {sk === 'documentLibraryCustom' && (
                    <Stack tokens={{ childrenGap: 10 }}>
                      <Dropdown
                        label="Biblioteca de documentos"
                        options={listDropdownOptionsWithLegacy(
                          attachmentLibraryDropdownOptions,
                          libBase.libraryTitle ?? ''
                        )}
                        selectedKey={(libBase.libraryTitle ?? '').trim() || ''}
                        onChange={(_, o) => {
                          if (!o) return;
                          onLinkedChildFormsChange(
                            patchChildById(linkedChildForms, cfg.id, {
                              childAttachmentLibrary: {
                                ...libBase,
                                libraryTitle: String(o.key),
                                sourceListLookupFieldInternalName: '',
                              },
                            })
                          );
                        }}
                        disabled={listsLoading}
                      />
                      {customLibLoading[cfg.id] === true && <Spinner label="A carregar campos da biblioteca…" />}
                      <Dropdown
                        label="Campo Lookup na biblioteca (referência à lista filha)"
                        options={
                          customLibFieldOpts[cfg.id] ?? [
                            {
                              key: '',
                              text:
                                customLibLoading[cfg.id] === true
                                  ? 'A carregar…'
                                  : '— selecione a biblioteca e a lista filha —',
                            },
                          ]
                        }
                        selectedKey={(libBase.sourceListLookupFieldInternalName ?? '').trim()}
                        onChange={(_, o) =>
                          o &&
                          onLinkedChildFormsChange(
                            patchChildById(linkedChildForms, cfg.id, {
                              childAttachmentLibrary: {
                                ...libBase,
                                sourceListLookupFieldInternalName: String(o.key),
                              },
                            })
                          )
                        }
                        disabled={customLibLoading[cfg.id] === true}
                      />
                      <FormManagerFolderTreeEditor
                        nodes={loadFolderTreeFromAttachmentLibrary(cfg.childAttachmentLibrary)}
                        onChange={(next) =>
                          onLinkedChildFormsChange(
                            patchChildById(linkedChildForms, cfg.id, {
                              childAttachmentLibrary: {
                                ...libBase,
                                folderTree: next,
                              },
                            })
                          )
                        }
                        disabled={listsLoading}
                        folderStepOptions={stepOpts}
                        showFolderStepPicker={stepOpts.length > 1}
                      />
                    </Stack>
                  )}
                      </Stack>
                    );
                  })()}
                </LinkedChildTabCollapseSection>
                <LinkedChildTabCollapseSection
                  title="Campos na etapa «Geral» (ordem)"
                  expanded={panels.fields}
                  onToggle={flip('fields')}
                  indent
                >
                  <Stack styles={{ root: { border: '1px solid #edebe9', padding: 8, borderRadius: 4 } }} tokens={{ childrenGap: 6 }}>
              {mainNames.length === 0 && (
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  Nenhum campo. Adicione abaixo.
                </Text>
              )}
              {mainNames.map((name, ni) => {
                const m = meta.find((x) => x.InternalName === name);
                return (
                  <Stack key={name} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    <Text variant="small" styles={{ root: { minWidth: 180 } }}>
                      {m ? `${m.Title} (${name})` : name}
                    </Text>
                    <IconButton
                      iconProps={{ iconName: 'Up' }}
                      title="Subir"
                      disabled={ni === 0}
                      onClick={() => {
                        const fn = mainNames.slice();
                        const t = fn[ni - 1];
                        fn[ni - 1] = fn[ni];
                        fn[ni] = t;
                        let next = setMainStepFieldNames(cfg, fn);
                        next = ensureFieldConfigsForNames(next, fn);
                        onLinkedChildFormsChange(
                          linkedChildForms.map((c) => (c.id === cfg.id ? next : c))
                        );
                      }}
                    />
                    <IconButton
                      iconProps={{ iconName: 'Down' }}
                      title="Descer"
                      disabled={ni === mainNames.length - 1}
                      onClick={() => {
                        const fn = mainNames.slice();
                        const t = fn[ni + 1];
                        fn[ni + 1] = fn[ni];
                        fn[ni] = t;
                        let next = setMainStepFieldNames(cfg, fn);
                        next = ensureFieldConfigsForNames(next, fn);
                        onLinkedChildFormsChange(
                          linkedChildForms.map((c) => (c.id === cfg.id ? next : c))
                        );
                      }}
                    />
                    <IconButton
                      iconProps={{ iconName: 'Delete' }}
                      title="Remover"
                      onClick={() => {
                        const fn = mainNames.filter((_, j) => j !== ni);
                        let next = setMainStepFieldNames(cfg, fn);
                        next = {
                          ...next,
                          fields: next.fields.filter((f) => fn.indexOf(f.internalName) !== -1),
                        };
                        onLinkedChildFormsChange(
                          linkedChildForms.map((c) => (c.id === cfg.id ? next : c))
                        );
                      }}
                    />
                    <DefaultButton
                      text="Regras…"
                      onClick={() => setLinkedFieldRuleTarget({ cfgId: cfg.id, internalName: name })}
                    />
                  </Stack>
                );
              })}
            </Stack>
            <Dropdown
              placeholder="Adicionar campo à etapa Geral…"
              options={addOptions.length ? addOptions : [{ key: '__none', text: '— sem campos disponíveis —' }]}
              selectedKey=""
              onChange={(_, o) => {
                if (!o || o.key === '__none') return;
                const key = String(o.key);
                if (mainNames.indexOf(key) !== -1) return;
                const fn = mainNames.concat([key]);
                let next = setMainStepFieldNames(cfg, fn);
                next = ensureFieldConfigsForNames(next, fn);
                onLinkedChildFormsChange(
                  linkedChildForms.map((c) => (c.id === cfg.id ? next : c))
                );
              }}
            />
                  </LinkedChildTabCollapseSection>
                  <LinkedChildTabCollapseSection
                    title="Regras condicionais"
                    expanded={panels.rules}
                    onToggle={flip('rules')}
                    indent
                  >
                    <FormManagerLinkedChildConditionalRulesBlock
                      rules={cfg.rules ?? []}
                      fieldOptions={meta
                        .filter(
                          (m) =>
                            isSelectableChildField(m) &&
                            m.InternalName !== cfg.parentLookupFieldInternalName.trim()
                        )
                        .map((m) => ({ key: m.InternalName, text: `${m.Title} (${m.InternalName})` }))}
                      onRulesChange={(next) =>
                        onLinkedChildFormsChange((prev) =>
                          prev.map((c) => (c.id === cfg.id ? { ...c, rules: next } : c))
                        )
                      }
                    />
                  </LinkedChildTabCollapseSection>
              </Stack>
            )}
          </Stack>
        );
      })}
    </Stack>
    {linkedFieldRuleTarget !== null &&
      ((): JSX.Element | null => {
        const lc = linkedChildForms.find((c) => c.id === linkedFieldRuleTarget.cfgId);
        if (!lc) return null;
        const name = linkedFieldRuleTarget.internalName;
        const childMeta = childMetaById[lc.id] ?? [];
        const fieldOpts: IDropdownOption[] = childMeta
          .filter(
            (m) =>
              isSelectableChildField(m) &&
              m.InternalName !== lc.parentLookupFieldInternalName.trim()
          )
          .map((m) => ({ key: m.InternalName, text: `${m.Title} (${m.InternalName})` }));
        const fc =
          lc.fields.find((f) => f.internalName === name) ??
          ({ internalName: name, sectionId: 'main' } as IFormFieldConfig);
        const fieldMeta = childMeta.find((x) => x.InternalName === name);
        const cfgId = lc.id;
        return (
          <FormFieldRulesPanel
            isOpen
            internalName={name}
            fieldConfig={fc}
            meta={fieldMeta}
            rules={lc.rules ?? []}
            fieldOptions={fieldOpts.length ? fieldOpts : [{ key: 'Title', text: 'Title' }]}
            onDismiss={() => setLinkedFieldRuleTarget(null)}
            onApply={(nextFc, editor) => {
              onLinkedChildFormsChange((prev) =>
                prev.map((c) => {
                  if (c.id !== cfgId) return c;
                  const rulesNext = mergeFieldRules(
                    c.rules ?? [],
                    name,
                    buildFieldUiRules(name, editor, nextFc)
                  );
                  const has = c.fields.some((f) => f.internalName === name);
                  const fieldsNext = has
                    ? c.fields.map((f) => (f.internalName === name ? { ...f, ...nextFc } : f))
                    : c.fields.concat([{ ...nextFc, internalName: name, sectionId: nextFc.sectionId ?? 'main' }]);
                  return { ...c, fields: fieldsNext, rules: rulesNext };
                })
              );
              setLinkedFieldRuleTarget(null);
            }}
          />
        );
      })()}
    </>
  );
}
