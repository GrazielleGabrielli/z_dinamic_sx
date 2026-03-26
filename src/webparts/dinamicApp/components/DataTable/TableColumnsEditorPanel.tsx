import * as React from 'react';
import { useState, useEffect, useMemo, useRef } from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  Checkbox,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  Separator,
  ChoiceGroup,
  IChoiceGroupOption,
  IconButton,
  Callout,
  Link,
  TooltipHost,
  Icon,
  Pivot,
  PivotItem,
} from '@fluentui/react';
import { FieldsService } from '../../../../services';
import type { IFieldMetadata } from '../../../../services';
import type {
  IListViewConfig,
  IListViewColumnConfig,
  IListViewModeConfig,
  IListViewFilterConfig,
  IPaginationConfig,
  IPdfTemplateConfig,
  ITableLayoutCssSlots,
  ITableRowStyleRule,
  TTableRowRuleOperator,
  TPaginationLayout,
  TFilterOperator,
} from '../../core/config/types';
import { PdfTemplateEditor } from './PdfTemplateEditor';
import {
  DINAMIC_SX_TABLE_CLASS,
  TABLE_LAYOUT_EDITOR_GROUPS,
  TABLE_LAYOUT_EDITOR_ROWS,
  TABLE_LAYOUT_SLOT_ORDER,
} from './tableLayoutClasses';
import { toTableRowRuleDataToken } from '../../core/table/utils/tableRowStyleRuleEval';
import { TableLayoutSlotPreview } from './TableLayoutSlotPreview';

interface ITableColumnsEditorPanelProps {
  isOpen: boolean;
  listTitle: string;
  listView: IListViewConfig;
  pagination: IPaginationConfig;
  pdfTemplate?: IPdfTemplateConfig;
  onSave: (listView: IListViewConfig, pagination: IPaginationConfig, pdfTemplate?: IPdfTemplateConfig) => void;
  onDismiss: () => void;
}

interface IFieldOption {
  meta: IFieldMetadata;
  selected: boolean;
  label: string;
  expandField: string;
}

const EXPANDABLE = ['lookup', 'lookupmulti', 'user', 'usermulti'];

const SIMPLE_FIELD_TYPES = ['text', 'multiline', 'number', 'currency', 'boolean', 'choice', 'multichoice', 'datetime', 'url'];

const USER_EXPAND_FIELDS: IDropdownOption[] = [
  { key: 'Id', text: 'Id' },
  { key: 'Title', text: 'Title' },
  { key: 'EMail', text: 'EMail' },
  { key: 'LoginName', text: 'LoginName' },
];

function toFieldOption(meta: IFieldMetadata, existing?: IListViewColumnConfig): IFieldOption {
  const selected = existing !== undefined;
  const needsExpand = EXPANDABLE.indexOf(meta.MappedType) !== -1;
  const expandField = needsExpand
    ? (existing?.expandField ?? meta.LookupField ?? 'Title')
    : '';
  return {
    meta,
    selected,
    label: existing?.label ?? meta.Title,
    expandField,
  };
}

function buildOptions(
  fields: IFieldMetadata[],
  currentColumns: IListViewColumnConfig[]
): IFieldOption[] {
  const byName = new Map(fields.map((f) => [f.InternalName, f]));
  const selectedSet = new Set(currentColumns.map((c) => c.field));
  const ordered: IFieldOption[] = [];
  currentColumns.forEach((c) => {
    const meta = byName.get(c.field);
    if (meta) ordered.push(toFieldOption(meta, c));
  });
  fields.forEach((f) => {
    if (!selectedSet.has(f.InternalName)) ordered.push(toFieldOption(f, undefined));
  });
  return ordered;
}

function buildExpandOptionsFromLookupList(fields: IFieldMetadata[]): IDropdownOption[] {
  const simple = fields.filter(
    (f) =>
      SIMPLE_FIELD_TYPES.indexOf(f.MappedType) !== -1 &&
      f.InternalName !== 'Id' &&
      f.InternalName !== 'Title'
  );
  const options: IDropdownOption[] = [
    { key: 'Id', text: 'Id' },
    { key: 'Title', text: 'Title' },
  ];
  simple.forEach((f) => options.push({ key: f.InternalName, text: `${f.Title} (${f.InternalName})` }));
  return options;
}

const PAGE_SIZE_OPTIONS = [5, 10, 20, 50, 100];

const PAGINATION_LAYOUT_OPTIONS: IChoiceGroupOption[] = [
  { key: 'buttons', text: 'Botões (Anterior / Próxima)' },
  { key: 'numbered', text: 'Numerada (Página X + botões)' },
  { key: 'compact', text: 'Compacta (linha única)' },
  { key: 'paged', text: 'Páginas (1 … 4 5 6 …)' },
];

const ROW_STYLE_RULE_OPERATORS: { key: TTableRowRuleOperator; text: string }[] = [
  { key: 'eq', text: 'É igual a' },
  { key: 'ne', text: 'É diferente de' },
  { key: 'contains', text: 'Contém' },
  { key: 'startsWith', text: 'Começa com' },
  { key: 'endsWith', text: 'Termina com' },
  { key: 'empty', text: 'Está vazio' },
  { key: 'notEmpty', text: 'Não está vazio' },
];

const VIEW_MODE_OPERATORS: IDropdownOption[] = [
  { key: 'eq', text: 'Igual a' },
  { key: 'ne', text: 'Diferente de' },
  { key: 'contains', text: 'Contém' },
  { key: 'gt', text: 'Maior que' },
  { key: 'ge', text: 'Maior ou igual' },
  { key: 'lt', text: 'Menor que' },
  { key: 'le', text: 'Menor ou igual' },
];

const FORMULA_TOKENS: { token: string; label: string }[] = [
  { token: '[me]', label: 'Usuário atual (ID)' },
  { token: '[myId]', label: 'Meu ID' },
  { token: '[myName]', label: 'Meu nome' },
  { token: '[myEmail]', label: 'Meu e-mail' },
  { token: '[myLogin]', label: 'Meu login' },
  { token: '[myDepartment]', label: 'Meu departamento' },
  { token: '[myJobTitle]', label: 'Meu cargo' },
  { token: '[today]', label: 'Data de hoje' },
  { token: '[now]', label: 'Data e hora atuais' },
  { token: '[tomorrow]', label: 'Amanhã' },
  { token: '[yesterday]', label: 'Ontem' },
  { token: '[startOfMonth]', label: 'Início do mês' },
  { token: '[endOfMonth]', label: 'Fim do mês' },
  { token: '[startOfYear]', label: 'Início do ano' },
  { token: '[endOfYear]', label: 'Fim do ano' },
  { token: '[siteTitle]', label: 'Título do site' },
  { token: '[siteUrl]', label: 'URL do site' },
  { token: '[listTitle]', label: 'Título da lista' },
  { token: '[empty]', label: 'Vazio' },
  { token: '[null]', label: 'Null' },
  { token: '[true]', label: 'Verdadeiro' },
  { token: '[false]', label: 'Falso' },
];

const DEFAULT_VIEW_MODES_FALLBACK: IListViewModeConfig[] = [
  { id: 'all', label: 'Todas', filters: [] },
  { id: 'mine', label: 'Minhas', filters: [{ field: 'Author/Id', operator: 'eq', value: '[Me]' }] },
];

function viewModeFilterSummary(filters: IListViewFilterConfig[]): string {
  if (!filters || filters.length === 0) return 'Sem filtros';
  return filters.map((f) => `${f.field} ${f.operator} "${f.value}"`).join(' e ');
}

export const TableColumnsEditorPanel: React.FC<ITableColumnsEditorPanelProps> = ({
  isOpen,
  listTitle,
  listView,
  pagination,
  pdfTemplate,
  onSave,
  onDismiss,
}) => {
  const [activeTab, setActiveTab] = useState<string>('lista');
  const [layoutSubTab, setLayoutSubTab] = useState<string>('geral');
  const [localPdfTemplate, setLocalPdfTemplate] = useState<IPdfTemplateConfig | undefined>(pdfTemplate);
  const [loading, setLoading] = useState(false);
  const [options, setOptions] = useState<IFieldOption[]>([]);
  const [lookupListFields, setLookupListFields] = useState<Record<string, IFieldMetadata[]>>({});
  const [paginationEnabled, setPaginationEnabled] = useState(pagination.enabled);
  const [pageSize, setPageSize] = useState(pagination.pageSize);
  const [paginationLayout, setPaginationLayout] = useState<TPaginationLayout>(pagination.layout ?? 'buttons');
  const [pdfExportEnabled, setPdfExportEnabled] = useState(listView.pdfExportEnabled ?? false);
  const [cssSlotsState, setCssSlotsState] = useState<ITableLayoutCssSlots>(() => ({
    ...(listView.customTableCssSlots ?? {}),
  }));
  const [customTableCssExtra, setCustomTableCssExtra] = useState(listView.customTableCss ?? '');
  const [rowStyleRules, setRowStyleRules] = useState<ITableRowStyleRule[]>(() => [
    ...(listView.tableRowStyleRules ?? []),
  ]);
  const [viewModes, setViewModes] = useState<IListViewModeConfig[]>(
    listView.viewModes?.length ? listView.viewModes : DEFAULT_VIEW_MODES_FALLBACK
  );
  const [activeViewModeId, setActiveViewModeId] = useState<string>(listView.activeViewModeId ?? 'all');
  const [viewModeEditingId, setViewModeEditingId] = useState<string | null>(null);
  const [viewModeEditLabel, setViewModeEditLabel] = useState('');
  const [viewModeEditFilters, setViewModeEditFilters] = useState<IListViewFilterConfig[]>([]);
  const [formulasFilterIndex, setFormulasFilterIndex] = useState<number | null>(null);
  const [formulasTarget, setFormulasTarget] = useState<HTMLElement | null>(null);
  const panelWasOpenRef = useRef(false);

  const fieldsService = useMemo(() => new FieldsService(), []);
  const layoutRowBySlot = useMemo(
    () => new Map(TABLE_LAYOUT_EDITOR_ROWS.map((r) => [r.slot, r] as const)),
    []
  );

  useEffect(() => {
    if (!isOpen || !listTitle.trim()) return;
    setLoading(true);
    setLookupListFields({});
    fieldsService
      .getVisibleFields(listTitle.trim())
      .then((f) => {
        const configured = listView.columns ?? [];
        const effectiveColumns =
          configured.length === 0 && f.some((field) => field.InternalName === 'Title')
            ? [{ field: 'Title' }]
            : configured;
        setOptions(buildOptions(f, effectiveColumns));
        const listIds = f
          .filter((x) => EXPANDABLE.indexOf(x.MappedType) !== -1 && x.LookupList)
          .map((x) => x.LookupList as string);
        const uniqueIds = listIds.filter((id, i) => listIds.indexOf(id) === i);
        return Promise.all(
          uniqueIds.map((id) =>
            fieldsService.getFields(id).then((fields) => ({ id, fields }))
          )
        );
      })
      .then((results) => {
        const next: Record<string, IFieldMetadata[]> = {};
        results.forEach(({ id, fields }) => { next[id] = fields; });
        setLookupListFields((prev) => ({ ...prev, ...next }));
      })
      .then(() => setLoading(false), () => setLoading(false));
  }, [isOpen, listTitle]);

  useEffect(() => {
    if (!isOpen) {
      panelWasOpenRef.current = false;
      return;
    }
    if (panelWasOpenRef.current) {
      return;
    }
    panelWasOpenRef.current = true;
    setPaginationEnabled(pagination.enabled);
    setPageSize(pagination.pageSize);
    setPaginationLayout(pagination.layout ?? 'buttons');
    setViewModes(listView.viewModes?.length ? listView.viewModes : DEFAULT_VIEW_MODES_FALLBACK);
    setActiveViewModeId(listView.activeViewModeId ?? 'all');
    setLocalPdfTemplate(pdfTemplate);
    setPdfExportEnabled(listView.pdfExportEnabled ?? false);
    setCssSlotsState({ ...(listView.customTableCssSlots ?? {}) });
    setCustomTableCssExtra(listView.customTableCss ?? '');
    setRowStyleRules([...(listView.tableRowStyleRules ?? [])]);
    setLayoutSubTab('geral');
  }, [isOpen, listView, pagination, pdfTemplate]);


  const toggle = (internalName: string): void => {
    setOptions((prev) =>
      prev.map((o) => (o.meta.InternalName === internalName ? { ...o, selected: !o.selected } : o))
    );
  };

  const setLabel = (internalName: string, label: string): void => {
    setOptions((prev) =>
      prev.map((o) => (o.meta.InternalName === internalName ? { ...o, label } : o))
    );
  };

  const setExpandField = (internalName: string, expandField: string): void => {
    setOptions((prev) =>
      prev.map((o) => (o.meta.InternalName === internalName ? { ...o, expandField } : o))
    );
  };

  const getExpandFieldOptions = (meta: IFieldMetadata): IDropdownOption[] => {
    if (meta.MappedType === 'user' || meta.MappedType === 'usermulti') return USER_EXPAND_FIELDS;
    if (meta.LookupList && lookupListFields[meta.LookupList]) {
      return buildExpandOptionsFromLookupList(lookupListFields[meta.LookupList]);
    }
    return [{ key: 'Title', text: 'Title' }, { key: 'Id', text: 'Id' }];
  };

  const filterFieldOptions = useMemo((): IDropdownOption[] => {
    const empty: IDropdownOption = { key: '', text: '— selecione —' };
    const rest: IDropdownOption[] = [];
    for (let i = 0; i < options.length; i++) {
      const o = options[i];
      if (EXPANDABLE.indexOf(o.meta.MappedType) === -1) {
        rest.push({ key: o.meta.InternalName, text: `${o.meta.Title} (${o.meta.InternalName})` });
      } else {
        const expandOpts = o.meta.MappedType === 'user' || o.meta.MappedType === 'usermulti'
          ? USER_EXPAND_FIELDS
          : (o.meta.LookupList && lookupListFields[o.meta.LookupList])
            ? buildExpandOptionsFromLookupList(lookupListFields[o.meta.LookupList])
            : [{ key: 'Title', text: 'Title' }, { key: 'Id', text: 'Id' }];
        for (let j = 0; j < expandOpts.length; j++) {
          const opt = expandOpts[j];
          rest.push({
            key: `${o.meta.InternalName}/${String(opt.key)}`,
            text: `${o.meta.Title} – ${opt.text}`,
          });
        }
      }
    }
    return [empty, ...rest];
  }, [options, lookupListFields]);

  const rowRuleFieldOptions: IDropdownOption[] = useMemo(() => {
    const empty: IDropdownOption = { key: '', text: '— selecione o campo —' };
    return [
      empty,
      ...options.map((o) => ({
        key: o.meta.InternalName,
        text: `${o.meta.Title} (${o.meta.InternalName})`,
      })),
    ];
  }, [options]);

  const addRowStyleRule = (): void => {
    setRowStyleRules((prev) => [
      ...prev,
      {
        id: `r_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
        field: '',
        operator: 'eq',
        value: '',
        rowCss: 'background: #fffbeb !important;',
      },
    ]);
  };

  const updateRowStyleRule = (index: number, part: Partial<ITableRowStyleRule>): void => {
    setRowStyleRules((prev) => {
      const next = prev.slice();
      if (next[index]) next[index] = { ...next[index], ...part };
      return next;
    });
  };

  const removeRowStyleRule = (index: number): void => {
    setRowStyleRules((prev) => prev.filter((_, i) => i !== index));
  };

  const handleSave = (): void => {
    const columns: IListViewColumnConfig[] = options
      .filter((o) => o.selected)
      .map((o) => {
        const base = {
          field: o.meta.InternalName,
          label: o.label.trim() ? o.label : o.meta.Title,
        };
        if (EXPANDABLE.indexOf(o.meta.MappedType) !== -1) {
          return { ...base, expandField: o.expandField.trim() || 'Title' };
        }
        return base;
      });
    const nextPagination: IPaginationConfig = {
      ...pagination,
      enabled: paginationEnabled,
      pageSize,
      layout: paginationLayout,
      pageSizeOptions: pagination.pageSizeOptions?.length ? pagination.pageSizeOptions : PAGE_SIZE_OPTIONS,
    };
    const nextSlots: ITableLayoutCssSlots = {};
    for (let i = 0; i < TABLE_LAYOUT_SLOT_ORDER.length; i++) {
      const slot = TABLE_LAYOUT_SLOT_ORDER[i];
      const t = (cssSlotsState[slot] ?? '').trim();
      if (t) nextSlots[slot] = t;
    }
    const extraTrim = customTableCssExtra.trim();
    const nextRowRules: ITableRowStyleRule[] = [];
    const seenIds = new Set<string>();
    for (let i = 0; i < rowStyleRules.length; i++) {
      const r = rowStyleRules[i];
      const field = r.field.trim();
      const rowCss = r.rowCss.trim();
      if (!field || !rowCss) continue;
      let id = toTableRowRuleDataToken(r.id.trim() || `r_${Date.now()}_${i}`);
      while (seenIds.has(id)) {
        id = toTableRowRuleDataToken(`${id}_${i}`);
      }
      seenIds.add(id);
      nextRowRules.push({
        id,
        field,
        operator: r.operator,
        value: r.value,
        rowCss,
      });
    }
    onSave(
      {
        ...listView,
        columns,
        viewModes,
        activeViewModeId,
        pdfExportEnabled,
        ...(Object.keys(nextSlots).length > 0 ? { customTableCssSlots: nextSlots } : { customTableCssSlots: undefined }),
        ...(extraTrim ? { customTableCss: extraTrim } : { customTableCss: undefined }),
        ...(nextRowRules.length > 0 ? { tableRowStyleRules: nextRowRules } : { tableRowStyleRules: undefined }),
      },
      nextPagination,
      localPdfTemplate
    );
    onDismiss();
  };

  const viewModeDefaultOptions: IDropdownOption[] = viewModes.map((m) => ({ key: m.id, text: m.label }));
  const startViewModeAdd = (): void => {
    setViewModeEditLabel('Novo modo');
    setViewModeEditFilters([]);
    setViewModeEditingId(`custom_${Date.now()}`);
  };
  const startViewModeEdit = (m: IListViewModeConfig): void => {
    setViewModeEditingId(m.id);
    setViewModeEditLabel(m.label);
    setViewModeEditFilters(m.filters?.length ? m.filters.slice() : []);
  };
  const saveViewModeEdit = (): void => {
    if (viewModeEditingId === null) return;
    const next = viewModes.slice();
    let idx = -1;
    for (let i = 0; i < next.length; i++) { if (next[i].id === viewModeEditingId) { idx = i; break; } }
    const entry: IListViewModeConfig = { id: viewModeEditingId, label: viewModeEditLabel.trim() || viewModeEditingId, filters: viewModeEditFilters };
    if (idx >= 0) next[idx] = entry;
    else next.push(entry);
    setViewModes(next);
    setViewModeEditingId(null);
  };
  const removeViewMode = (id: string): void => {
    if (id === 'all' || id === 'mine') return;
    const next = viewModes.filter((m) => m.id !== id);
    setViewModes(next);
    if (activeViewModeId === id) setActiveViewModeId(next[0]?.id ?? 'all');
  };
  const addViewModeFilter = (): void => setViewModeEditFilters([...viewModeEditFilters, { field: '', operator: 'eq', value: '' }]);
  const removeViewModeFilter = (i: number): void => setViewModeEditFilters(viewModeEditFilters.filter((_, idx) => idx !== i));
  const updateViewModeFilter = (i: number, part: Partial<IListViewFilterConfig>): void => {
    const next = viewModeEditFilters.slice();
    next[i] = { ...next[i], ...part };
    setViewModeEditFilters(next);
  };

  const openFormulas = (index: number, target: HTMLElement): void => {
    setFormulasFilterIndex(index);
    setFormulasTarget(target);
  };

  const applyFormula = (index: number, token: string): void => {
    updateViewModeFilter(index, { value: token });
    setFormulasFilterIndex(null);
    setFormulasTarget(null);
  };

  const pdfFieldOptions: IDropdownOption[] = useMemo(
    () => [
      { key: '', text: '— inserir campo —' },
      ...options.map((o) => ({
        key: o.meta.InternalName,
        text: `${o.meta.Title} (${o.meta.InternalName})`,
      })),
    ],
    [options]
  );

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.custom}
      customWidth="98vw"
      styles={{
        main: { width: 'min(98vw, calc(100vw - 16px))', maxWidth: 'min(98vw, calc(100vw - 16px))' },
        scrollableContent: { overflowX: 'hidden' },
        content: { overflowX: 'hidden', minWidth: 0 },
      }}
      headerText="Editar lista / tabela"
      closeButtonAriaLabel="Fechar"
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="Salvar" onClick={handleSave} />
          <DefaultButton text="Cancelar" onClick={onDismiss} />
        </Stack>
      )}
    >
      <div style={{ paddingTop: 16, minWidth: 0, maxWidth: '100%', boxSizing: 'border-box' }}>
        {loading ? (
          <Stack horizontalAlign="center" tokens={{ childrenGap: 12 }} style={{ padding: 24 }}>
            <Spinner size={SpinnerSize.medium} />
            <Text variant="small">Carregando campos...</Text>
          </Stack>
        ) : (
          <Stack tokens={{ childrenGap: 16 }} styles={{ root: { minWidth: 0, maxWidth: '100%' } }}>
            <Pivot
              selectedKey={activeTab}
              onLinkClick={(item) => item?.props?.itemKey !== undefined && item?.props?.itemKey !== null && setActiveTab(String(item.props.itemKey))}
              styles={{ root: { marginBottom: 8, flexWrap: 'wrap', maxWidth: '100%' } }}
            >
              <PivotItem itemKey="lista" headerText="Lista">
                <Stack tokens={{ childrenGap: 16 }} styles={{ root: { paddingTop: 8, minWidth: 0, maxWidth: '100%' } }}>
            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                Paginação
              </Text>
              <Checkbox
                label="Habilitar paginação"
                checked={paginationEnabled}
                onChange={(_, v) => setPaginationEnabled(!!v)}
              />
              {paginationEnabled && (
                <>
                  <Dropdown
                    label="Itens por página"
                    selectedKey={String(pageSize)}
                    options={(pagination.pageSizeOptions?.length ? pagination.pageSizeOptions : PAGE_SIZE_OPTIONS).map(
                      (n) => ({ key: String(n), text: String(n) })
                    )}
                    onChange={(_, opt) => setPageSize(Number(opt?.key) || 10)}
                    styles={{ root: { maxWidth: 120 } }}
                  />
                  <ChoiceGroup
                    label="Layout da paginação"
                    options={PAGINATION_LAYOUT_OPTIONS}
                    selectedKey={paginationLayout}
                    onChange={(_, opt) => opt && setPaginationLayout(opt.key as TPaginationLayout)}
                  />
                </>
              )}
            </Stack>

            <Separator />

            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                Modos de visualização
              </Text>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Ex.: Todas (sem filtro), Minhas (Author/Id eq [Me]), ou filtros customizados. O usuário alterna entre eles na lista.
              </Text>
              <Dropdown
                label="Modo padrão"
                options={viewModeDefaultOptions}
                selectedKey={activeViewModeId}
                onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => opt && setActiveViewModeId(String(opt.key))}
                styles={{ root: { maxWidth: 280 } }}
              />
              {viewModes.map((m) => (
                <div key={m.id} style={{ padding: 10, border: '1px solid #edebe9', borderRadius: 6, background: viewModeEditingId === m.id ? '#f3f9ff' : '#fff' }}>
                  {viewModeEditingId === m.id ? (
                    <Stack tokens={{ childrenGap: 10 }}>
                      <TextField label="Nome" value={viewModeEditLabel} onChange={(_: React.FormEvent, v?: string) => setViewModeEditLabel(v ?? '')} />
                      {viewModeEditFilters.map((f, i) => (
                        <Stack key={i} horizontal wrap verticalAlign="start" tokens={{ childrenGap: 8 }} styles={{ root: { width: '100%', minWidth: 0 } }}>
                          <Dropdown placeholder="Campo" options={filterFieldOptions} selectedKey={f.field || ''} onChange={(_: React.FormEvent, opt?: IDropdownOption) => updateViewModeFilter(i, { field: (opt?.key as string) ?? '' })} styles={{ root: { flex: '1 1 200px', minWidth: 0, maxWidth: '100%' } }} />
                          <Dropdown options={VIEW_MODE_OPERATORS} selectedKey={f.operator} onChange={(_: React.FormEvent, opt?: IDropdownOption) => opt !== undefined && opt !== null && updateViewModeFilter(i, { operator: String(opt.key) as TFilterOperator })} styles={{ root: { flex: '0 0 auto', width: 120, minWidth: 100 } }} />
                          <Stack tokens={{ childrenGap: 2 }} styles={{ root: { flex: '1 1 180px', minWidth: 0, maxWidth: '100%' } }}>
                            <TextField placeholder="Valor ou [Me]" value={f.value} onChange={(_: React.FormEvent, v?: string) => updateViewModeFilter(i, { value: v ?? '' })} />
                            <Text variant="small" styles={{ root: { marginTop: 0 } }}>
                              <Link onClick={(e) => openFormulas(i, e.currentTarget as HTMLElement)} style={{ cursor: 'pointer', color: '#0078d4' }}>Fórmulas</Link>
                            </Text>
                          </Stack>
                          <IconButton iconProps={{ iconName: 'Delete' }} title="Remover filtro" onClick={() => removeViewModeFilter(i)} />
                        </Stack>
                      ))}
                      <DefaultButton text="Adicionar filtro" onClick={addViewModeFilter} />
                      <Stack horizontal tokens={{ childrenGap: 8 }}>
                        <PrimaryButton text="Salvar" onClick={saveViewModeEdit} />
                        <DefaultButton text="Cancelar" onClick={() => setViewModeEditingId(null)} />
                      </Stack>
                    </Stack>
                  ) : (
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                      <Stack tokens={{ childrenGap: 2 }}>
                        <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>{m.label}</Text>
                        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{viewModeFilterSummary(m.filters)}</Text>
                      </Stack>
                      <Stack horizontal tokens={{ childrenGap: 4 }}>
                        <IconButton iconProps={{ iconName: 'Edit' }} title="Editar" onClick={() => startViewModeEdit(m)} />
                        <IconButton iconProps={{ iconName: 'Delete' }} title="Remover" onClick={() => removeViewMode(m.id)} disabled={m.id === 'all' || m.id === 'mine'} />
                      </Stack>
                    </Stack>
                  )}
                </div>
              ))}
              {viewModeEditingId !== null && !viewModes.some((m) => m.id === viewModeEditingId) && (
                <div style={{ padding: 10, border: '1px solid #c7e0f4', borderRadius: 6, background: '#f3f9ff' }}>
                  <Stack tokens={{ childrenGap: 10 }}>
                    <TextField label="Nome" value={viewModeEditLabel} onChange={(_: React.FormEvent, v?: string) => setViewModeEditLabel(v ?? '')} />
                    {viewModeEditFilters.map((f, i) => (
                      <Stack key={i} horizontal wrap verticalAlign="start" tokens={{ childrenGap: 8 }} styles={{ root: { width: '100%', minWidth: 0 } }}>
                        <Dropdown placeholder="Campo" options={filterFieldOptions} selectedKey={f.field || ''} onChange={(_: React.FormEvent, opt?: IDropdownOption) => updateViewModeFilter(i, { field: (opt?.key as string) ?? '' })} styles={{ root: { flex: '1 1 200px', minWidth: 0, maxWidth: '100%' } }} />
                        <Dropdown options={VIEW_MODE_OPERATORS} selectedKey={f.operator} onChange={(_: React.FormEvent, opt?: IDropdownOption) => opt !== undefined && opt !== null && updateViewModeFilter(i, { operator: String(opt.key) as TFilterOperator })} styles={{ root: { flex: '0 0 auto', width: 120, minWidth: 100 } }} />
                        <Stack tokens={{ childrenGap: 2 }} styles={{ root: { flex: '1 1 180px', minWidth: 0, maxWidth: '100%' } }}>
                          <TextField placeholder="Valor ou [Me]" value={f.value} onChange={(_: React.FormEvent, v?: string) => updateViewModeFilter(i, { value: v ?? '' })} />
                          <Text variant="small" styles={{ root: { marginTop: 0 } }}>
                            <Link onClick={(e) => openFormulas(i, e.currentTarget as HTMLElement)} style={{ cursor: 'pointer', color: '#0078d4' }}>Fórmulas</Link>
                          </Text>
                        </Stack>
                        <IconButton iconProps={{ iconName: 'Delete' }} title="Remover filtro" onClick={() => removeViewModeFilter(i)} />
                      </Stack>
                    ))}
                    <DefaultButton text="Adicionar filtro" onClick={addViewModeFilter} />
                    <Stack horizontal tokens={{ childrenGap: 8 }}>
                      <PrimaryButton text="Adicionar modo" onClick={saveViewModeEdit} />
                      <DefaultButton text="Cancelar" onClick={() => setViewModeEditingId(null)} />
                    </Stack>
                  </Stack>
                </div>
              )}
              {viewModeEditingId === null && <DefaultButton text="Adicionar modo de visualização" onClick={startViewModeAdd} />}
            </Stack>

            <Separator />

            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                Colunas
              </Text>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Marque as colunas que deseja exibir. Para lookups e usuários, escolha o campo de exibição.
              </Text>
            </Stack>
            {options.map((o) => (
              <Stack
                key={o.meta.InternalName}
                horizontal
                wrap
                tokens={{ childrenGap: 12 }}
                verticalAlign="start"
                styles={{ root: { padding: '8px 0', borderBottom: '1px solid #f3f2f1', width: '100%', minWidth: 0 } }}
              >
                <Checkbox
                  checked={o.selected}
                  onChange={() => toggle(o.meta.InternalName)}
                  ariaLabel={o.meta.Title}
                  styles={{ root: { flex: '0 0 auto', marginTop: 4 } }}
                />
                <Stack tokens={{ childrenGap: 4 }} styles={{ root: { flex: '1 1 240px', minWidth: 0, maxWidth: '100%' } }}>
                  <Text variant="small" styles={{ root: { color: '#605e5c', marginBottom: 2, wordBreak: 'break-word' } }}>
                    Campo: {o.meta.InternalName}
                  </Text>
                  <TextField
                    label="Rótulo (cabeçalho na tabela)"
                    value={o.label}
                    onChange={(_, v) => setLabel(o.meta.InternalName, v ?? '')}
                    disabled={!o.selected}
                    placeholder={o.meta.Title}
                    styles={{ root: { maxWidth: '100%' } }}
                  />
                  {EXPANDABLE.indexOf(o.meta.MappedType) !== -1 && o.selected && (
                    <Dropdown
                      label="Campo expandido (lookup/user)"
                      selectedKey={o.expandField || 'Title'}
                      options={getExpandFieldOptions(o.meta)}
                      onChange={(_, opt) => setExpandField(o.meta.InternalName, (opt?.key as string) ?? 'Title')}
                      styles={{ root: { maxWidth: '100%' } }}
                    />
                  )}
                </Stack>
                <Text variant="small" styles={{ root: { color: '#a19f9d', flex: '0 0 auto', maxWidth: '100%', wordBreak: 'break-word' } }}>
                  {o.meta.MappedType}
                </Text>
              </Stack>
            ))}
                </Stack>
              </PivotItem>
              <PivotItem itemKey="pdf" headerText="PDF">
                <Stack tokens={{ childrenGap: 12 }} styles={{ root: { paddingTop: 8, minWidth: 0, maxWidth: '100%' } }}>
                  <Checkbox
                    label="Exibir botão Exportar PDF ao lado do seletor de abas"
                    checked={pdfExportEnabled}
                    onChange={(_, v) => setPdfExportEnabled(!!v)}
                  />
                  <PdfTemplateEditor
                    value={localPdfTemplate}
                    onChange={setLocalPdfTemplate}
                    fieldOptions={pdfFieldOptions}
                  />
                </Stack>
              </PivotItem>
              <PivotItem itemKey="excel" headerText="Excel">
                <Stack tokens={{ childrenGap: 8 }} styles={{ root: { paddingTop: 16, minWidth: 0, maxWidth: '100%' } }}>
                  <Text variant="medium" styles={{ root: { color: '#605e5c' } }}>
                    Exportação para Excel em breve.
                  </Text>
                </Stack>
              </PivotItem>
              <PivotItem itemKey="layout" headerText="Layout">
                <Pivot
                  selectedKey={layoutSubTab}
                  onLinkClick={(item) =>
                    item?.props?.itemKey !== undefined &&
                    item?.props?.itemKey !== null &&
                    setLayoutSubTab(String(item.props.itemKey))
                  }
                  styles={{ root: { marginBottom: 4, flexWrap: 'wrap', maxWidth: '100%' } }}
                >
                  <PivotItem itemKey="geral" headerText="Geral">
                <Stack tokens={{ childrenGap: 0 }} styles={{ root: { paddingTop: 4, minWidth: 0, maxWidth: '100%', paddingBottom: 24 } }}>
                  <Stack
                    styles={{
                      root: {
                        position: 'sticky',
                        top: 0,
                        zIndex: 2,
                        background: '#ffffff',
                        paddingBottom: 12,
                        paddingTop: 4,
                        marginBottom: 4,
                        borderBottom: '1px solid #edebe9',
                      },
                    }}
                  >
                    <Text variant="small" styles={{ root: { color: '#323130', lineHeight: 1.55 } }}>
                      <strong>Como usar:</strong> em cada campo, só declarações CSS (propriedade: valor;), aplicadas à classe indicada.
                      Colunas específicas com <span style={{ fontFamily: 'monospace' }}>[data-field=&quot;NomeInterno&quot;]</span> vão no bloco
                      &quot;CSS adicional&quot; no final. Use <span style={{ fontFamily: 'monospace' }}>!important</span> se o estilo padrão da lista
                      tiver prioridade.
                    </Text>
                  </Stack>
                  {TABLE_LAYOUT_EDITOR_GROUPS.map((group, gi) => (
                    <Stack key={group.id} tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: gi === 0 ? 8 : 28 } }}>
                      <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 10 }} styles={{ root: { flexWrap: 'wrap' } }}>
                        <div style={{ width: 4, minHeight: 28, background: '#0078d4', borderRadius: 2, flexShrink: 0, marginTop: 2 }} />
                        <Stack tokens={{ childrenGap: 4 }} styles={{ root: { flex: '1 1 240px', minWidth: 0 } }}>
                          <Text variant="medium" styles={{ root: { fontWeight: 600, color: '#201f1e' } }}>
                            {group.label}
                          </Text>
                          <Text variant="small" styles={{ root: { color: '#605e5c', lineHeight: 1.45 } }}>
                            {group.blurb}
                          </Text>
                        </Stack>
                      </Stack>
                      {group.slots.map((slotKey) => {
                        const row = layoutRowBySlot.get(slotKey);
                        if (!row) return null;
                        const cls = DINAMIC_SX_TABLE_CLASS[row.slot];
                        return (
                          <Stack
                            key={row.slot}
                            horizontal
                            wrap
                            verticalAlign="start"
                            tokens={{ childrenGap: 16 }}
                            styles={{
                              root: {
                                padding: '14px 16px',
                                background: '#faf9f8',
                                borderRadius: 8,
                                border: '1px solid #edebe9',
                                borderLeftWidth: 3,
                                borderLeftColor: '#0078d4',
                              },
                            }}
                          >
                            <Stack styles={{ root: { flex: '2 1 340px', minWidth: 0 } }} tokens={{ childrenGap: 6 }}>
                              <Text variant="smallPlus" styles={{ root: { fontWeight: 600, color: '#201f1e' } }}>
                                {row.title}
                              </Text>
                              <Text variant="small" styles={{ root: { color: '#605e5c', lineHeight: 1.45 } }}>
                                {row.hint}
                              </Text>
                              <Text
                                variant="small"
                                styles={{
                                  root: {
                                    fontFamily: 'monospace',
                                    color: '#0078d4',
                                    marginTop: 2,
                                    wordBreak: 'break-all',
                                  },
                                }}
                              >
                                .{cls}
                              </Text>
                              <TextField
                                multiline
                                resizable
                                rows={3}
                                ariaLabel={`CSS para ${row.title}`}
                                value={cssSlotsState[row.slot] ?? ''}
                                onChange={(_, v) =>
                                  setCssSlotsState((prev) => ({
                                    ...prev,
                                    [row.slot]: v ?? '',
                                  }))
                                }
                                placeholder="ex.: background: #f3f2f1; font-weight: 600;"
                                styles={{ root: { marginTop: 4, maxWidth: '100%' } }}
                              />
                            </Stack>
                            <Stack styles={{ root: { flex: '1 1 240px', minWidth: 200, maxWidth: '100%' } }}>
                              <TableLayoutSlotPreview
                                slot={row.slot}
                                cssBody={cssSlotsState[row.slot] ?? ''}
                                variant="embedded"
                              />
                            </Stack>
                          </Stack>
                        );
                      })}
                    </Stack>
                  ))}
                  <Separator styles={{ root: { marginTop: 28, marginBottom: 4 } }} />
                  <Stack
                    tokens={{ childrenGap: 10 }}
                    styles={{
                      root: {
                        padding: 16,
                        background: '#fff9f5',
                        borderRadius: 8,
                        border: '1px solid #edebe9',
                        borderLeft: '3px solid #ca5010',
                      },
                    }}
                  >
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                      <Icon iconName="Code" styles={{ root: { color: '#ca5010', fontSize: 18 } }} />
                      <Text variant="medium" styles={{ root: { fontWeight: 600, color: '#201f1e' } }}>
                        CSS adicional (regras livres)
                      </Text>
                    </Stack>
                    <Text variant="small" styles={{ root: { color: '#605e5c', lineHeight: 1.45 } }}>
                      Vários seletores, <span style={{ fontFamily: 'monospace' }}>:hover</span>, media queries. É aplicado depois dos blocos por
                      componente.
                    </Text>
                    <TextField
                      label="CSS livre"
                      multiline
                      rows={6}
                      value={customTableCssExtra}
                      onChange={(_, v) => setCustomTableCssExtra(v ?? '')}
                      placeholder={
                        `.${DINAMIC_SX_TABLE_CLASS.row}:nth-child(even) { background: #faf9f8 !important; }\n` +
                        `.${DINAMIC_SX_TABLE_CLASS.cell}[data-field="Title"] { font-weight: 600; }`
                      }
                      styles={{ root: { maxWidth: '100%' } }}
                    />
                  </Stack>
                </Stack>
                  </PivotItem>
                  <PivotItem itemKey="regras" headerText="Regras">
                    <Stack tokens={{ childrenGap: 14 }} styles={{ root: { paddingTop: 8, paddingBottom: 24, minWidth: 0, maxWidth: '100%' } }}>
                      <Text variant="small" styles={{ root: { color: '#605e5c', lineHeight: 1.55 } }}>
                        Aplique CSS na <strong>linha inteira</strong> (<span style={{ fontFamily: 'monospace' }}>&lt;tr&gt;</span>) quando o valor
                        exibido de um campo atender à condição. A comparação usa o mesmo texto que aparece na célula (incluindo lookups).
                        Várias regras podem valer ao mesmo tempo; cada uma adiciona um marcador em{' '}
                        <span style={{ fontFamily: 'monospace' }}>data-dinamic-rules</span>.
                      </Text>
                      {rowStyleRules.map((rule, ri) => {
                        const valueDisabled = rule.operator === 'empty' || rule.operator === 'notEmpty';
                        return (
                          <Stack
                            key={rule.id}
                            tokens={{ childrenGap: 10 }}
                            styles={{
                              root: {
                                padding: 14,
                                background: '#f3f9ff',
                                borderRadius: 8,
                                border: '1px solid #c7e0f4',
                                borderLeftWidth: 3,
                                borderLeftColor: '#0078d4',
                              },
                            }}
                          >
                            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                              <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                                Regra {ri + 1}
                              </Text>
                              <IconButton iconProps={{ iconName: 'Delete' }} title="Remover regra" onClick={() => removeRowStyleRule(ri)} />
                            </Stack>
                            <Dropdown
                              label="Campo"
                              selectedKey={rule.field || ''}
                              options={rowRuleFieldOptions}
                              onChange={(_: React.FormEvent, opt?: IDropdownOption) =>
                                updateRowStyleRule(ri, { field: String(opt?.key ?? '') })
                              }
                              styles={{ root: { maxWidth: '100%' } }}
                            />
                            <Dropdown
                              label="Condição"
                              selectedKey={rule.operator}
                              options={ROW_STYLE_RULE_OPERATORS.map((o) => ({ key: o.key, text: o.text }))}
                              onChange={(_: React.FormEvent, opt?: IDropdownOption) =>
                                opt && updateRowStyleRule(ri, { operator: String(opt.key) as TTableRowRuleOperator })
                              }
                              styles={{ root: { maxWidth: '100%' } }}
                            />
                            <TextField
                              label="Valor"
                              value={rule.value}
                              onChange={(_, v) => updateRowStyleRule(ri, { value: v ?? '' })}
                              disabled={valueDisabled}
                              description={valueDisabled ? 'Não usado para “vazio” / “não vazio”.' : undefined}
                            />
                            <TextField
                              label="CSS da linha (declarações)"
                              multiline
                              rows={3}
                              value={rule.rowCss}
                              onChange={(_, v) => updateRowStyleRule(ri, { rowCss: v ?? '' })}
                              placeholder="ex.: background: #fef9c3 !important; font-weight: 600;"
                              styles={{ root: { maxWidth: '100%' } }}
                            />
                            <Text variant="small" styles={{ root: { fontFamily: 'monospace', color: '#605e5c' } }}>
                              Marcador: data-dinamic-rules~=&quot;{toTableRowRuleDataToken(rule.id)}&quot;
                            </Text>
                          </Stack>
                        );
                      })}
                      <DefaultButton text="Adicionar regra" iconProps={{ iconName: 'Add' }} onClick={addRowStyleRule} />
                    </Stack>
                  </PivotItem>
                </Pivot>
              </PivotItem>
            </Pivot>
          </Stack>
        )}
        {formulasTarget && formulasFilterIndex !== null && (
          <Callout
            target={formulasTarget}
            onDismiss={() => { setFormulasFilterIndex(null); setFormulasTarget(null); }}
            role="dialog"
            ariaLabel="Fórmulas disponíveis"
          >
            <Stack tokens={{ childrenGap: 4 }} styles={{ root: { padding: 12, minWidth: 260, maxHeight: 320, overflowY: 'auto' } }}>
              <Text variant="medium" styles={{ root: { fontWeight: 600, marginBottom: 4 } }}>Fórmulas</Text>
              <Text variant="small" styles={{ root: { color: '#605e5c', marginBottom: 8 } }}>Clique para inserir no valor do filtro.</Text>
              {FORMULA_TOKENS.map((item) => (
                <Stack key={item.token} horizontal horizontalAlign="space-between" verticalAlign="center" styles={{ root: { padding: '6px 0', borderBottom: '1px solid #edebe9', cursor: 'pointer' } }} onClick={() => applyFormula(formulasFilterIndex, item.token)}>
                  <Text variant="small" styles={{ root: { fontFamily: 'monospace' } }}>{item.token}</Text>
                  <Text variant="small" styles={{ root: { color: '#605e5c', flex: 1, marginLeft: 8 } }}>{item.label}</Text>
                  <TooltipHost content={item.label} calloutProps={{ gapSpace: 4 }}>
                    <Icon iconName="Unknown" style={{ fontSize: 12, color: '#605e5c', marginLeft: 4 }} onClick={(e) => e.stopPropagation()} />
                  </TooltipHost>
                </Stack>
              ))}
              <Text variant="small" styles={{ root: { marginTop: 8, color: '#605e5c' } }}>Query: use [query:nome] para parâmetro da URL.</Text>
            </Stack>
          </Callout>
        )}
      </div>
    </Panel>
  );
};
