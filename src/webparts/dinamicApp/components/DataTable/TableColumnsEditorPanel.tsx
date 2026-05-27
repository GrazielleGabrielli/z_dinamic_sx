import * as React from 'react';
import { useState, useEffect, useMemo, useRef, useCallback } from 'react';
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
  Toggle,
  MessageBar,
  MessageBarType,
  Pivot,
  PivotItem,
} from '@fluentui/react';
import { FieldsService, SYSTEM_METADATA_FIELDS } from '../../../../services';
import type { IFieldMetadata } from '../../../../services';
import type {
  IProjectManagementColumnConfig,
  IProjectManagementConfig,
  IProjectManagementRuleConfig,
  IListViewConfig,
  IListViewColumnConfig,
  IListViewModeConfig,
  IListViewModeAccessConfig,
  IListViewModeDefaultRule,
  IListViewFilterConfig,
  ITableFilterFieldConfig,
  IListViewChromeButtonConfig,
  IListViewSortConfig,
  IPaginationConfig,
  IPdfTemplateConfig,
  IListRowActionConfig,
  IListRowActionFieldRule,
  IListRowActionVisibility,
  ITableRowStyleRule,
  TListRowActionFieldRuleOp,
  TListRowActionIconPreset,
  TTableRowRuleOperator,
  TPaginationLayout,
  TFilterOperator,
  TViewMode,
  TListViewDisplayMode,
  TViewModePicker,
  TListViewChromeButtonSlot,
  TListViewColumnBreakpoint,
} from '../../core/config/types';
import { PdfTemplateEditor } from './PdfTemplateEditor';
import {
  DINAMIC_SX_TABLE_CLASS,
  TABLE_LAYOUT_EDITOR_ROWS,
  mergeCustomTableCss,
  mergeRowStyleRulesCss,
  DEFAULT_TABLE_LAYOUT_CSS,
} from './tableLayoutClasses';
import { isNoteFieldMeta } from '../../core/listView';
import { toTableRowRuleDataToken } from '../../core/table/utils/tableRowStyleRuleEval';
import { TableLayoutLivePreview } from './TableLayoutLivePreview';
import { VIEW_MODE_PICKER_OPTIONS, normalizeViewModePicker } from './viewModePickerLayouts';
import {
  sanitizeListTableEditorBundle,
  sanitizeListViewChromeButtons,
} from '../../core/config/validators';
import { ViewModeAccessSection, accessSummary } from '../shared/ViewModeAccessSection';
import {
  LIST_VIEW_COLUMN_BREAKPOINT_LABEL,
  LIST_VIEW_COLUMN_BREAKPOINT_MIN_PX,
  LIST_VIEW_COLUMN_BREAKPOINT_ORDER,
  clampListColumnSpan,
  resolveListColumnSpanAtBreakpoint,
} from '../../core/listView/listViewColumnBreakpoints';

interface ITableColumnsEditorPanelProps {
  isOpen: boolean;
  mode: TViewMode;
  listTitle: string;
  listWebServerRelativeUrl?: string;
  /** Site da página do web part (grupos para permissão de modo). */
  pageWebServerRelativeUrl: string;
  listView: IListViewConfig;
  pagination: IPaginationConfig;
  projectManagement?: IProjectManagementConfig;
  pdfTemplate?: IPdfTemplateConfig;
  onSave: (
    listView: IListViewConfig,
    pagination: IPaginationConfig,
    pdfTemplate?: IPdfTemplateConfig,
    projectManagement?: IProjectManagementConfig
  ) => void;
  onDismiss: () => void;
}

interface IFieldOption {
  meta: IFieldMetadata;
  selected: boolean;
  label: string;
  expandField: string;
  expandFieldsSelected: string[];
  columnSpanByBreakpoint: Partial<Record<TListViewColumnBreakpoint, number>>;
}

const EXPANDABLE = ['lookup', 'lookupmulti', 'user', 'usermulti'];

type TTableFilterDragZone = 'fixed' | 'advanced';

const LIST_TABLE_COL_DND = 'dinamicSx:listTableCol:';

const LIST_TABLE_FILTER_DND = 'dinamicSx:listTableFilter:';

function parseTableFilterDragPayload(raw: string): {
  fromZone: TTableFilterDragZone;
  field: string;
  fromIdx: number;
} | null {
  if (!raw.startsWith(LIST_TABLE_FILTER_DND)) return null;
  const rest = raw.slice(LIST_TABLE_FILTER_DND.length);
  const parts = rest.split('\t');
  if (parts.length < 3) return null;
  const zoneChar = parts[0];
  const field = parts[1];
  const fromIdx = parseInt(parts[2], 10);
  if (!field || Number.isNaN(fromIdx)) return null;
  const fromZone: TTableFilterDragZone = zoneChar === 'a' ? 'advanced' : 'fixed';
  return { fromZone, field, fromIdx };
}

const SIMPLE_FIELD_TYPES = ['text', 'multiline', 'number', 'currency', 'boolean', 'choice', 'multichoice', 'datetime', 'url'];

function getTableFilterFieldKey(o: IFieldOption): string {
  return EXPANDABLE.indexOf(o.meta.MappedType) !== -1 && o.meta.LookupField
    ? `${o.meta.InternalName}/${o.meta.MappedType === 'user' || o.meta.MappedType === 'usermulti' ? 'Title' : o.meta.LookupField}`
    : o.meta.InternalName;
}

function tableFilterEntryMatchesOption(entry: ITableFilterFieldConfig, o: IFieldOption): boolean {
  const fk = getTableFilterFieldKey(o);
  return entry.field === fk || entry.field === o.meta.InternalName;
}

function findOptionForTableFilterEntry(entry: ITableFilterFieldConfig, opts: IFieldOption[]): IFieldOption | undefined {
  return opts.find((o) => tableFilterEntryMatchesOption(entry, o));
}

function partitionTableFilterFields(fields: ITableFilterFieldConfig[]): {
  fixed: ITableFilterFieldConfig[];
  advanced: ITableFilterFieldConfig[];
} {
  const fixed: ITableFilterFieldConfig[] = [];
  const advanced: ITableFilterFieldConfig[] = [];
  for (let i = 0; i < fields.length; i++) {
    const f = fields[i];
    if (f.placement === 'advanced') advanced.push(f);
    else fixed.push(f);
  }
  return { fixed, advanced };
}

function mergeTableFilterPartitions(fixed: ITableFilterFieldConfig[], advanced: ITableFilterFieldConfig[]): ITableFilterFieldConfig[] {
  return [...fixed, ...advanced];
}

function normalizeTableFilterFieldsOrderLocal(fields: ITableFilterFieldConfig[]): ITableFilterFieldConfig[] {
  const { fixed, advanced } = partitionTableFilterFields(fields);
  return mergeTableFilterPartitions(fixed, advanced);
}

const USER_EXPAND_FIELDS: IDropdownOption[] = [
  { key: 'Id', text: 'Id' },
  { key: 'Title', text: 'Title' },
  { key: 'EMail', text: 'EMail' },
  { key: 'LoginName', text: 'LoginName' },
];

function sanitizeColumnSpanMap(
  map: Partial<Record<TListViewColumnBreakpoint, number>> | undefined
): Partial<Record<TListViewColumnBreakpoint, number>> | undefined {
  if (!map) return undefined;
  const out: Partial<Record<TListViewColumnBreakpoint, number>> = {};
  for (let i = 0; i < LIST_VIEW_COLUMN_BREAKPOINT_ORDER.length; i++) {
    const k = LIST_VIEW_COLUMN_BREAKPOINT_ORDER[i];
    const v = map[k];
    if (v !== undefined) out[k] = clampListColumnSpan(v);
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

function mergeColumnSpanFromGroup(group: IListViewColumnConfig[]): Partial<Record<TListViewColumnBreakpoint, number>> {
  for (let i = 0; i < group.length; i++) {
    const m = group[i].columnSpanByBreakpoint;
    if (m && Object.keys(m).length > 0) {
      const s = sanitizeColumnSpanMap(m);
      if (s) return s;
    }
  }
  return {};
}

function toFieldOption(meta: IFieldMetadata, existing?: IListViewColumnConfig): IFieldOption {
  const selected = existing !== undefined;
  const needsExpand = EXPANDABLE.indexOf(meta.MappedType) !== -1;
  const ef = (existing?.expandField ?? meta.LookupField ?? 'Title').trim() || 'Title';
  const expandFieldsSelected = needsExpand && existing ? [ef] : needsExpand ? [] : [];
  return {
    meta,
    selected,
    label: existing?.label ?? meta.Title,
    expandField: needsExpand ? (existing ? ef : meta.LookupField ?? 'Title') : '',
    expandFieldsSelected,
    columnSpanByBreakpoint: mergeColumnSpanFromGroup(existing ? [existing] : []),
  };
}

function fieldOptionFromColumnGroup(meta: IFieldMetadata, group: IListViewColumnConfig[]): IFieldOption {
  const needsExpand = EXPANDABLE.indexOf(meta.MappedType) !== -1;
  const keys = needsExpand
    ? Array.from(new Set(group.map((c) => (c.expandField ?? meta.LookupField ?? 'Title').trim() || 'Title')))
    : [];
  const first = group[0];
  return {
    meta,
    selected: true,
    label: first?.label?.trim() ? first.label : meta.Title,
    expandField: keys[0] ?? 'Title',
    expandFieldsSelected: keys,
    columnSpanByBreakpoint: mergeColumnSpanFromGroup(group),
  };
}

function applyColumnsToOptions(opts: IFieldOption[], cols: IListViewColumnConfig[]): IFieldOption[] {
  const map = new Map(opts.map((o) => [o.meta.InternalName, o]));
  const byField = new Map<string, IListViewColumnConfig[]>();
  for (let i = 0; i < cols.length; i++) {
    const c = cols[i];
    if (!byField.has(c.field)) byField.set(c.field, []);
    byField.get(c.field)!.push(c);
  }
  const ordered: IFieldOption[] = [];
  const seen = new Set<string>();
  for (let i = 0; i < cols.length; i++) {
    const c = cols[i];
    if (seen.has(c.field)) continue;
    seen.add(c.field);
    const o = map.get(c.field);
    if (!o) continue;
    const group = byField.get(c.field) ?? [c];
    ordered.push(fieldOptionFromColumnGroup(o.meta, group));
    map.delete(c.field);
  }
  map.forEach((o) =>
    ordered.push({
      ...o,
      selected: false,
      expandFieldsSelected: EXPANDABLE.indexOf(o.meta.MappedType) !== -1 ? [] : [],
      columnSpanByBreakpoint: {},
    })
  );
  return ordered;
}

function buildOptions(
  fields: IFieldMetadata[],
  currentColumns: IListViewColumnConfig[]
): IFieldOption[] {
  const byName = new Map(fields.map((f) => [f.InternalName, f]));
  const byField = new Map<string, IListViewColumnConfig[]>();
  for (let i = 0; i < currentColumns.length; i++) {
    const c = currentColumns[i];
    if (!byField.has(c.field)) byField.set(c.field, []);
    byField.get(c.field)!.push(c);
  }
  const ordered: IFieldOption[] = [];
  const used = new Set<string>();
  currentColumns.forEach((c) => {
    if (used.has(c.field)) return;
    used.add(c.field);
    const meta = byName.get(c.field);
    if (meta) ordered.push(fieldOptionFromColumnGroup(meta, byField.get(c.field) ?? [c]));
  });
  fields.forEach((f) => {
    if (used.has(f.InternalName)) return;
    ordered.push(toFieldOption(f, undefined));
  });
  return ordered;
}

function buildExpandOptionsFromLookupList(fields: IFieldMetadata[]): IDropdownOption[] {
  const simple = fields.filter(
    (f) =>
      SIMPLE_FIELD_TYPES.indexOf(f.MappedType) !== -1 &&
      f.InternalName !== 'Id' &&
      f.InternalName !== 'Title' &&
      !isNoteFieldMeta(f)
  );
  const options: IDropdownOption[] = [
    { key: 'Id', text: 'Id' },
    { key: 'Title', text: 'Title' },
  ];
  simple.forEach((f) => options.push({ key: f.InternalName, text: `${f.Title} (${f.InternalName})` }));
  return options;
}

const PAGE_SIZE_OPTIONS = [5, 10, 20, 50, 100];

const LIST_COLUMN_SPAN_OPTIONS: IDropdownOption[] = Array.from({ length: 12 }, (_, i) => ({
  key: i + 1,
  text: String(i + 1),
}));

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

const LIST_ROW_ICON_PRESET_OPTIONS: IDropdownOption[] = [
  { key: 'view', text: 'Ver (olho)' },
  { key: 'edit', text: 'Editar' },
  { key: 'link', text: 'Link' },
  { key: 'custom', text: 'Personalizado (nome Fluent)' },
];

const LIST_ROW_SCOPE_OPTIONS: IChoiceGroupOption[] = [
  { key: 'icon', text: 'Somente ícone' },
  { key: 'wholeRow', text: 'Linha ou card inteiro' },
];

const ROW_ACTION_FIELD_RULE_OP_OPTIONS: IDropdownOption[] = [
  { key: 'eq', text: 'Igual a (=)' },
  { key: 'ne', text: 'Diferente de (≠)' },
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

const VIEW_MODE_PICKER_CHOICES: IChoiceGroupOption[] = VIEW_MODE_PICKER_OPTIONS.map((o) => ({
  key: o.key,
  text: o.text,
}));

const DEFAULT_PROJECT_COLUMNS: IProjectManagementColumnConfig[] = [];

const LIST_VIEW_CHROME_SLOT_OPTIONS: IDropdownOption[] = [
  { key: 'toolbarAfterViewMode', text: 'Barra: após modo de visualização' },
  { key: 'toolbarAfterTableCardsToggle', text: 'Barra: após alternância Tabela/Cards' },
  { key: 'toolbarAfterPdfExport', text: 'Barra: após Exportar PDF' },
  { key: 'toolbarBeforeClearFilters', text: 'Barra: antes de Remover filtros' },
  { key: 'filtersAfterAdvancedToggle', text: 'Filtros: após botão dos filtros avançados' },
  { key: 'filtersBelowControls', text: 'Filtros: linha abaixo dos controles' },
];

const LIST_CHROME_ACTION_OPTIONS: IDropdownOption[] = [
  { key: 'redirect', text: 'Abrir URL' },
  { key: 'reload', text: 'Recarregar página' },
];

const LIST_SORT_DIRECTION_OPTIONS: IChoiceGroupOption[] = [
  { key: 'asc', text: 'Ascendente (A–Z, mais antigo primeiro)' },
  { key: 'desc', text: 'Descendente (Z–A, mais recente primeiro)' },
];

/** OData nos itens da lista; disponível sempre na ordenação, mesmo sem coluna visível. */
const LIST_SORT_SYSTEM_COLUMN_OPTIONS: IDropdownOption[] = [
  { key: 'Id', text: 'ID do item (Id)' },
  { key: 'Created', text: 'Criado (Created)' },
  { key: 'Modified', text: 'Modificado (Modified)' },
];

function createDefaultListChromeButton(): IListViewChromeButtonConfig {
  return {
    id: `chrome_${Date.now()}`,
    label: 'Novo botão',
    actionKind: 'redirect',
    url: 'https://',
    variant: 'default',
    slot: 'toolbarAfterViewMode',
  };
}
function normalizeHexColor(input: string | undefined, fallback: string): string {
  const raw = (input ?? '').trim();
  return /^#([0-9a-fA-F]{6})$/.test(raw) ? raw : fallback;
}

function viewModeFilterSummary(filters: IListViewFilterConfig[]): string {
  if (!filters || filters.length === 0) return 'Sem filtros';
  return filters.map((f) => `${f.field} ${f.operator} "${f.value}"`).join(' e ');
}

type TListTabListaSection =
  | 'pagination'
  | 'tableSort'
  | 'listTableCards'
  | 'viewModes'
  | 'columns'
  | 'filterFields'
  | 'chromeButtons'
  | 'rowActions'
  | 'exportPdfExcel';

function ListTabListaCollapse(props: {
  title: string;
  isOpen: boolean;
  onToggle: () => void;
  children: React.ReactNode;
  zoneDrop?: {
    highlight: boolean;
    onDragOver: (e: React.DragEvent) => void;
    onDragLeave?: (e: React.DragEvent) => void;
    onDrop: (e: React.DragEvent) => void;
  };
}): JSX.Element {
  const zd = props.zoneDrop;
  const zoneBorder = zd?.highlight ? '2px solid #0078d4' : undefined;
  const zoneBg = zd?.highlight ? 'rgba(0, 120, 212, 0.06)' : undefined;
  return (
    <Stack
      styles={{
        root: {
          border: '1px solid #e8e6e4',
          borderRadius: 12,
          background: '#ffffff',
          boxShadow: '0 1px 3px rgba(0, 0, 0, 0.06), 0 4px 14px rgba(0, 0, 0, 0.04)',
          overflow: 'hidden',
          maxWidth: '100%',
          minWidth: 0,
          width: '100%',
          boxSizing: 'border-box',
        },
      }}
    >
      <Stack
        horizontal
        verticalAlign="center"
        tokens={{ childrenGap: 2 }}
        onDragOver={zd?.onDragOver}
        onDragLeave={zd?.onDragLeave}
        onDrop={zd?.onDrop}
        styles={{
          root: {
            padding: '12px 14px',
            background: zoneBg ?? (props.isOpen ? '#f8f7f6' : '#ffffff'),
            borderBottom: props.isOpen ? '1px solid #edebe9' : undefined,
            userSelect: 'none',
            borderTop: !props.isOpen && zd?.highlight ? zoneBorder : undefined,
            boxSizing: 'border-box',
          },
        }}
      >
        <IconButton
          iconProps={{ iconName: props.isOpen ? 'ChevronDown' : 'ChevronRight' }}
          title={props.isOpen ? 'Recolher' : 'Expandir'}
          aria-expanded={props.isOpen}
          onClick={(e) => {
            e.preventDefault();
            props.onToggle();
          }}
          styles={{ root: { width: 32, height: 32 } }}
        />
        <Text
          variant="smallPlus"
          styles={{ root: { fontWeight: 600, cursor: 'pointer', flex: 1, color: '#323130' } }}
          onClick={props.onToggle}
        >
          {props.title}
        </Text>
      </Stack>
      {props.isOpen ? (
        <div
          onDragOver={zd?.onDragOver}
          onDragLeave={zd?.onDragLeave}
          onDrop={zd?.onDrop}
          style={{
            padding: '14px 14px 16px 18px',
            maxWidth: '100%',
            minWidth: 0,
            width: '100%',
            boxSizing: 'border-box',
            display: 'flex',
            flexDirection: 'column',
            gap: 12,
            borderTop: props.isOpen && zd?.highlight ? zoneBorder : undefined,
            backgroundColor: zoneBg ?? undefined,
          }}
        >
          {props.children}
        </div>
      ) : null}
    </Stack>
  );
}

export const TableColumnsEditorPanel: React.FC<ITableColumnsEditorPanelProps> = ({
  isOpen,
  mode,
  listTitle,
  listWebServerRelativeUrl,
  pageWebServerRelativeUrl,
  listView,
  pagination,
  projectManagement,
  pdfTemplate,
  onSave,
  onDismiss,
}) => {
  const lw = listWebServerRelativeUrl?.trim() || undefined;
  const [layoutSectionOpen, setLayoutSectionOpen] = useState<Partial<Record<'tableCss' | 'rowRules' | 'cardCss' | 'filterCss' | 'viewModeCss', boolean>>>({});
  const [listEditorPivotKey, setListEditorPivotKey] = useState<string>('geral');
  const [localPdfTemplate, setLocalPdfTemplate] = useState<IPdfTemplateConfig | undefined>(pdfTemplate);
  const [loading, setLoading] = useState(false);
  const [options, setOptions] = useState<IFieldOption[]>([]);
  const [lookupListFields, setLookupListFields] = useState<Record<string, IFieldMetadata[]>>({});
  const [paginationEnabled, setPaginationEnabled] = useState(pagination.enabled);
  const [pageSize, setPageSize] = useState(pagination.pageSize);
  const [paginationLayout, setPaginationLayout] = useState<TPaginationLayout>(pagination.layout ?? 'buttons');
  const [pdfExportEnabled, setPdfExportEnabled] = useState(listView.pdfExportEnabled ?? false);
  const [listCardViewEnabled, setListCardViewEnabled] = useState(listView.listCardViewEnabled ?? false);
  const [listDefaultDisplayMode, setListDefaultDisplayMode] = useState<TListViewDisplayMode>(
    listView.listDefaultDisplayMode === 'cards' ? 'cards' : 'table'
  );
  const [layoutCssText, setLayoutCssText] = useState<string>(
    mergeCustomTableCss(listView.customTableCssSlots, listView.customTableCss)
  );
  const [cardCssText, setCardCssText] = useState<string>(listView.customCardCss ?? '');
  const [filterCssText, setFilterCssText] = useState<string>(listView.customFilterCss ?? '');
  const [viewModeCssText, setViewModeCssText] = useState<string>(listView.customViewModeCss ?? '');
  const [layoutColor, setLayoutColor] = useState<string>('#0078d4');
  const [cardColor, setCardColor] = useState<string>('#0078d4');
  const [filterColor, setFilterColor] = useState<string>('#0078d4');
  const [viewModeColor, setViewModeColor] = useState<string>('#0078d4');
  const [projectColumns, setProjectColumns] = useState<IProjectManagementColumnConfig[]>(
    projectManagement?.columns?.length ? projectManagement.columns : DEFAULT_PROJECT_COLUMNS
  );
  const [rowStyleRules, setRowStyleRules] = useState<ITableRowStyleRule[]>(() => [
    ...(listView.tableRowStyleRules ?? []),
  ]);
  const [rowActions, setRowActions] = useState<IListRowActionConfig[]>(() => [...(listView.listRowActions ?? [])]);
  const [visibilitySectionOpen, setVisibilitySectionOpen] = useState<Record<string, boolean>>({});
  const [tableFilterFields, setTableFilterFields] = useState<ITableFilterFieldConfig[]>(() =>
    normalizeTableFilterFieldsOrderLocal(listView.tableFilterFields?.slice() ?? [])
  );
  const [tableAdvancedFiltersTitle, setTableAdvancedFiltersTitle] = useState<string>(() =>
    listView.tableAdvancedFiltersTitle?.trim() ?? ''
  );
  const [listChromeButtons, setListChromeButtons] = useState<IListViewChromeButtonConfig[]>(() =>
    (listView.chromeButtons ?? []).map((b) => ({ ...b }))
  );
  const [listSortField, setListSortField] = useState<string>(() => listView.sort?.field?.trim() ?? '');
  const [listSortAscending, setListSortAscending] = useState<boolean>(() => listView.sort?.ascending !== false);
  const [filterEditorFixedOpen, setFilterEditorFixedOpen] = useState(true);
  const [filterEditorAdvancedOpen, setFilterEditorAdvancedOpen] = useState(true);
  const [tableFilterZoneDropHighlight, setTableFilterZoneDropHighlight] = useState<TTableFilterDragZone | null>(null);
  const layoutPreviewCss = useMemo(() => {
    const layout = layoutCssText.trim() || DEFAULT_TABLE_LAYOUT_CSS;
    const rules = mergeRowStyleRulesCss(rowStyleRules).trim();
    return [layout, rules].filter(Boolean).join('\n\n');
  }, [layoutCssText, rowStyleRules]);
  const layoutPreviewRuleTokens = useMemo(
    () =>
      rowStyleRules
        .filter((r) => (r.rowCss ?? '').trim().length > 0)
        .map((r) => toTableRowRuleDataToken(r.id))
        .slice(0, 2),
    [rowStyleRules]
  );
  const [ruleColorMap, setRuleColorMap] = useState<Record<string, string>>({});
  const [viewModes, setViewModes] = useState<IListViewModeConfig[]>(
    listView.viewModes?.length ? listView.viewModes : DEFAULT_VIEW_MODES_FALLBACK
  );
  const [activeViewModeId, setActiveViewModeId] = useState<string>(listView.activeViewModeId ?? 'all');
  const [listViewModePicker, setListViewModePicker] = useState<TViewModePicker>(
    normalizeViewModePicker(listView.viewModePicker)
  );
  const [viewModeDefaultRules, setViewModeDefaultRules] = useState<IListViewModeDefaultRule[]>(
    () => listView.viewModeDefaultRules?.map((r) => ({ ...r })) ?? []
  );
  const [viewModeEditingId, setViewModeEditingId] = useState<string | null>(null);
  const [viewModeEditLabel, setViewModeEditLabel] = useState('');
  const [viewModeEditFilters, setViewModeEditFilters] = useState<IListViewFilterConfig[]>([]);
  const [viewModeEditAccess, setViewModeEditAccess] = useState<IListViewModeAccessConfig | undefined>(undefined);
  const [formulasFilterIndex, setFormulasFilterIndex] = useState<number | null>(null);
  const [formulasTarget, setFormulasTarget] = useState<HTMLElement | null>(null);
  const panelWasOpenRef = useRef(false);
  const [carryListView, setCarryListView] = useState<IListViewConfig>(listView);
  const [jsonOpen, setJsonOpen] = useState(false);
  const [jsonPanelText, setJsonPanelText] = useState('');
  const [jsonPanelErr, setJsonPanelErr] = useState<string | undefined>(undefined);
  const [listTabListaSectionOpen, setListTabListaSectionOpen] = useState<
    Partial<Record<TListTabListaSection, boolean>>
  >({});

  const fieldsService = useMemo(() => new FieldsService(), []);

  useEffect(() => {
    if (!isOpen || !listTitle.trim()) return;
    setLoading(true);
    setLookupListFields({});
    fieldsService
      .getVisibleFields(listTitle.trim(), lw)
      .then((f) => {
        const extra = SYSTEM_METADATA_FIELDS.filter(
          (sf) => !f.some((x) => x.InternalName === sf.InternalName)
        );
        const allFields = [...f, ...extra];
        const configured = listView.columns ?? [];
        const effectiveColumns =
          configured.length === 0 && allFields.some((field) => field.InternalName === 'Title')
            ? [{ field: 'Title' }]
            : configured;
        setOptions(buildOptions(allFields, effectiveColumns));
        const listIds = f
          .filter((x) => EXPANDABLE.indexOf(x.MappedType) !== -1 && x.LookupList)
          .map((x) => x.LookupList as string);
        const uniqueIds = listIds.filter((id, i) => listIds.indexOf(id) === i);
        return Promise.all(
          uniqueIds.map((id) =>
            fieldsService.getFields(id, lw).then((fields) => ({ id, fields }))
          )
        );
      })
      .then((results) => {
        const next: Record<string, IFieldMetadata[]> = {};
        results.forEach(({ id, fields }) => { next[id] = fields; });
        setLookupListFields((prev) => ({ ...prev, ...next }));
      })
      .then(() => setLoading(false), () => setLoading(false));
  }, [isOpen, listTitle, lw]);

  useEffect(() => {
    if (!isOpen) {
      panelWasOpenRef.current = false;
      return;
    }
    if (panelWasOpenRef.current) {
      return;
    }
    panelWasOpenRef.current = true;
    setCarryListView(listView);
    setPaginationEnabled(pagination.enabled);
    setPageSize(pagination.pageSize);
    setPaginationLayout(pagination.layout ?? 'buttons');
    setViewModes(listView.viewModes?.length ? listView.viewModes : DEFAULT_VIEW_MODES_FALLBACK);
    setActiveViewModeId(listView.activeViewModeId ?? 'all');
    setViewModeDefaultRules(listView.viewModeDefaultRules?.map((r) => ({ ...r })) ?? []);
    setListViewModePicker(normalizeViewModePicker(listView.viewModePicker));
    setLocalPdfTemplate(pdfTemplate);
    setPdfExportEnabled(listView.pdfExportEnabled ?? false);
    setListCardViewEnabled(listView.listCardViewEnabled ?? false);
    setListDefaultDisplayMode(listView.listDefaultDisplayMode === 'cards' ? 'cards' : 'table');
    setLayoutCssText(mergeCustomTableCss(listView.customTableCssSlots, listView.customTableCss));
    setCardCssText(listView.customCardCss ?? '');
    setFilterCssText(listView.customFilterCss ?? '');
    setViewModeCssText(listView.customViewModeCss ?? '');
    setLayoutSectionOpen({});
    setProjectColumns(projectManagement?.columns?.length ? projectManagement.columns : DEFAULT_PROJECT_COLUMNS);
    setRowStyleRules([...(listView.tableRowStyleRules ?? [])]);
    setRowActions([...(listView.listRowActions ?? [])]);
    setTableFilterFields(normalizeTableFilterFieldsOrderLocal(listView.tableFilterFields?.slice() ?? []));
    setTableAdvancedFiltersTitle(listView.tableAdvancedFiltersTitle?.trim() ?? '');
    setListChromeButtons((listView.chromeButtons ?? []).map((b) => ({ ...b })));
    setListSortField(listView.sort?.field?.trim() ?? '');
    setListSortAscending(listView.sort?.ascending !== false);
    setFilterEditorFixedOpen(true);
    setFilterEditorAdvancedOpen(true);
    setRuleColorMap({});
    setLayoutSectionOpen({});
    setListTabListaSectionOpen({});
  }, [isOpen, listView, pagination, pdfTemplate, projectManagement]);

  const showPdfExcelConfigTabs = mode !== 'list';

  const toggle = (internalName: string): void => {
    setOptions((prev) =>
      prev.map((o) => {
        if (o.meta.InternalName !== internalName) return o;
        const nextSel = !o.selected;
        if (!nextSel) {
          return { ...o, selected: false, expandFieldsSelected: [] };
        }
        const needsExpand = EXPANDABLE.indexOf(o.meta.MappedType) !== -1;
        const efs =
          needsExpand && (o.expandFieldsSelected?.length ?? 0) === 0
            ? [(o.expandField ?? o.meta.LookupField ?? 'Title').trim() || 'Title']
            : o.expandFieldsSelected ?? [];
        const ef0 = (efs[0] ?? o.meta.LookupField ?? 'Title').trim() || 'Title';
        return { ...o, selected: true, expandFieldsSelected: efs.length ? efs : [ef0], expandField: ef0 };
      })
    );
  };

  const toggleLookupExpandField = (internalName: string, expandKey: string, checked: boolean): void => {
    setOptions((prev) =>
      prev.map((o) => {
        if (o.meta.InternalName !== internalName) return o;
        let next = [...(o.expandFieldsSelected ?? [])];
        if (checked) {
          if (next.indexOf(expandKey) === -1) next.push(expandKey);
        } else {
          next = next.filter((k) => k !== expandKey);
          if (next.length === 0) next = ['Title'];
        }
        const ef0 = (next[0] ?? 'Title').trim() || 'Title';
        return { ...o, expandFieldsSelected: next, expandField: ef0 };
      })
    );
  };

  const setLabel = (internalName: string, label: string): void => {
    setOptions((prev) =>
      prev.map((o) => (o.meta.InternalName === internalName ? { ...o, label } : o))
    );
  };

  const setExpandField = (internalName: string, expandField: string): void => {
    setOptions((prev) =>
      prev.map((o) =>
        o.meta.InternalName === internalName
          ? { ...o, expandField, expandFieldsSelected: [expandField] }
          : o
      )
    );
  };

  const [listColumnDragOverIndex, setListColumnDragOverIndex] = useState<number | null>(null);

  const onListTableColumnDragStart =
    (rowIndex: number, selected: boolean) =>
    (e: React.DragEvent): void => {
      if (!selected) {
        e.preventDefault();
        return;
      }
      e.dataTransfer.setData('text/plain', `${LIST_TABLE_COL_DND}${rowIndex}`);
      e.dataTransfer.effectAllowed = 'move';
    };

  const onListTableColumnDragOver = (e: React.DragEvent, rowIndex: number): void => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    setListColumnDragOverIndex(rowIndex);
  };

  const onListTableColumnDragLeave = (): void => {
    setListColumnDragOverIndex(null);
  };

  const onListTableColumnDrop =
    (dropIndex: number) =>
    (e: React.DragEvent): void => {
      e.preventDefault();
      setListColumnDragOverIndex(null);
      const raw = e.dataTransfer.getData('text/plain');
      if (!raw.startsWith(LIST_TABLE_COL_DND)) return;
      const from = parseInt(raw.slice(LIST_TABLE_COL_DND.length), 10);
      if (Number.isNaN(from) || from === dropIndex) return;
      setOptions((prev) => {
        if (!prev[from]?.selected) return prev;
        const next = prev.slice();
        const [moved] = next.splice(from, 1);
        next.splice(dropIndex, 0, moved);
        return next;
      });
    };

  const [tableFilterDragOver, setTableFilterDragOver] = useState<{ zone: TTableFilterDragZone; index: number } | null>(null);

  const clearTableFilterDnDUi = (): void => {
    setTableFilterDragOver(null);
    setTableFilterZoneDropHighlight(null);
  };

  const moveTableFilterToZone = (field: string, placement: TTableFilterDragZone): void => {
    setTableFilterFields((prev) => {
      const { fixed, advanced } = partitionTableFilterFields(prev);
      const entry = fixed.find((f) => f.field === field) ?? advanced.find((f) => f.field === field);
      if (!entry) return prev;
      const restFixed = fixed.filter((f) => f.field !== field);
      const restAdv = advanced.filter((f) => f.field !== field);
      const labelPart = entry.label?.trim() ? { label: entry.label.trim() } : {};
      if (placement === 'fixed') {
        return mergeTableFilterPartitions([...restFixed, { field: entry.field, ...labelPart }], restAdv);
      }
      return mergeTableFilterPartitions(restFixed, [...restAdv, { field: entry.field, ...labelPart, placement: 'advanced' }]);
    });
  };

  const onTableFilterZoneDragOver =
    (targetZone: TTableFilterDragZone) =>
    (e: React.DragEvent): void => {
      e.preventDefault();
      e.stopPropagation();
      e.dataTransfer.dropEffect = 'move';
      setTableFilterDragOver(null);
      setTableFilterZoneDropHighlight(targetZone);
    };

  const onTableFilterZoneDrop =
    (targetZone: TTableFilterDragZone) =>
    (e: React.DragEvent): void => {
      e.preventDefault();
      e.stopPropagation();
      clearTableFilterDnDUi();
      const parsed = parseTableFilterDragPayload(e.dataTransfer.getData('text/plain'));
      if (!parsed || parsed.fromZone === targetZone) return;
      moveTableFilterToZone(parsed.field, targetZone);
      if (targetZone === 'fixed') setFilterEditorFixedOpen(true);
      else setFilterEditorAdvancedOpen(true);
    };

  const onTableFilterDragStart =
    (zone: TTableFilterDragZone, entry: ITableFilterFieldConfig, localIdx: number) =>
    (e: React.DragEvent): void => {
      const z = zone === 'advanced' ? 'a' : 'f';
      e.dataTransfer.setData('text/plain', `${LIST_TABLE_FILTER_DND}${z}\t${entry.field}\t${localIdx}`);
      e.dataTransfer.effectAllowed = 'move';
    };

  const onTableFilterDragOverRow = (e: React.DragEvent, zone: TTableFilterDragZone, localIdx: number): void => {
    e.preventDefault();
    e.stopPropagation();
    e.dataTransfer.dropEffect = 'move';
    setTableFilterZoneDropHighlight(null);
    setTableFilterDragOver({ zone, index: localIdx });
  };

  const onTableFilterDropRow =
    (dropZone: TTableFilterDragZone, dropLocalIdx: number) =>
    (e: React.DragEvent): void => {
      e.preventDefault();
      e.stopPropagation();
      clearTableFilterDnDUi();
      const parsed = parseTableFilterDragPayload(e.dataTransfer.getData('text/plain'));
      if (!parsed || parsed.fromZone !== dropZone) return;
      const from = parsed.fromIdx;
      if (from === dropLocalIdx) return;
      setTableFilterFields((prev) => {
        const { fixed, advanced } = partitionTableFilterFields(prev);
        const arr = dropZone === 'fixed' ? fixed.slice() : advanced.slice();
        if (from < 0 || from >= arr.length || dropLocalIdx < 0 || dropLocalIdx >= arr.length) return prev;
        const nextArr = arr.slice();
        const [moved] = nextArr.splice(from, 1);
        nextArr.splice(dropLocalIdx, 0, moved);
        return dropZone === 'fixed' ? mergeTableFilterPartitions(nextArr, advanced) : mergeTableFilterPartitions(fixed, nextArr);
      });
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
      if (isNoteFieldMeta(o.meta)) continue;
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

  const filterFieldUnselectedOptions = useMemo(
    () => options.filter((o) => !tableFilterFields.some((f) => tableFilterEntryMatchesOption(f, o))),
    [options, tableFilterFields]
  );

  const tableFiltersPartitioned = useMemo(() => partitionTableFilterFields(tableFilterFields), [tableFilterFields]);

  const listSortFieldOptions = useMemo((): IDropdownOption[] => {
    const seen = new Set<string>();
    const out: IDropdownOption[] = [];
    for (let s = 0; s < LIST_SORT_SYSTEM_COLUMN_OPTIONS.length; s++) {
      const opt = LIST_SORT_SYSTEM_COLUMN_OPTIONS[s];
      seen.add(String(opt.key));
      out.push(opt);
    }
    for (let i = 0; i < options.length; i++) {
      const o = options[i];
      if (!o.selected) continue;
      if (isNoteFieldMeta(o.meta)) continue;
      if (o.meta.MappedType === 'lookupmulti' || o.meta.MappedType === 'usermulti') continue;
      if (o.meta.MappedType === 'multichoice') continue;
      if (EXPANDABLE.indexOf(o.meta.MappedType) !== -1) {
        const keys =
          o.meta.MappedType === 'lookup' || o.meta.MappedType === 'user'
            ? o.expandFieldsSelected.length > 0
              ? o.expandFieldsSelected
              : [(o.expandField ?? 'Title').trim() || 'Title']
            : [(o.expandField ?? 'Title').trim() || 'Title'];
        const expandOpts = getExpandFieldOptions(o.meta);
        const labelFor = (k: string): string => {
          const hit = expandOpts.find((x) => String(x.key) === k);
          return `${o.meta.Title} – ${hit?.text ?? k}`;
        };
        for (let j = 0; j < keys.length; j++) {
          const ek = (keys[j] ?? 'Title').trim() || 'Title';
          const path = `${o.meta.InternalName}/${ek}`;
          if (seen.has(path)) continue;
          seen.add(path);
          out.push({
            key: path,
            text: keys.length > 1 ? labelFor(ek) : o.label.trim() ? `${o.label} (${path})` : labelFor(ek),
          });
        }
      } else {
        const key = o.meta.InternalName;
        if (seen.has(key)) continue;
        seen.add(key);
        out.push({
          key,
          text: o.label.trim() ? `${o.label} (${key})` : `${o.meta.Title} (${key})`,
        });
      }
    }
    return out;
  }, [options, lookupListFields]);

  useEffect(() => {
    if (!listSortField.trim()) return;
    const keys = new Set(listSortFieldOptions.map((opt) => String(opt.key)));
    if (!keys.has(listSortField)) setListSortField('');
  }, [listSortField, listSortFieldOptions]);

  const renderConfiguredTableFilterRow = (
    entry: ITableFilterFieldConfig,
    zone: TTableFilterDragZone,
    localIdx: number
  ): React.ReactNode => {
    const o = findOptionForTableFilterEntry(entry, options);
    const defaultLabel = o?.meta.Title ?? entry.field;
    const dragHighlight =
      tableFilterDragOver?.zone === zone && tableFilterDragOver?.index === localIdx;
    return (
      <Stack
        key={`tf_${zone}_${entry.field}_${localIdx}`}
        horizontal
        wrap
        tokens={{ childrenGap: 12 }}
        verticalAlign="center"
        onDragOver={(e) => onTableFilterDragOverRow(e, zone, localIdx)}
        onDrop={onTableFilterDropRow(zone, localIdx)}
        onDragEnd={clearTableFilterDnDUi}
        styles={{
          root: {
            padding: '8px 0',
            borderBottom: '1px solid #f3f2f1',
            width: '100%',
            minWidth: 0,
            boxSizing: 'border-box',
            borderTop: dragHighlight ? '2px solid #0078d4' : undefined,
            backgroundColor: dragHighlight ? 'rgba(0, 120, 212, 0.04)' : undefined,
          },
        }}
      >
        <span
          draggable
          onDragStart={onTableFilterDragStart(zone, entry, localIdx)}
          onDragEnd={clearTableFilterDnDUi}
          title="Arrastar para reordenar dentro deste grupo; largue no outro collapse para mudar de zona"
          style={{
            cursor: 'grab',
            display: 'flex',
            alignItems: 'center',
            flexShrink: 0,
            minWidth: 22,
          }}
        >
          <Icon iconName="GripperBarVertical" styles={{ root: { fontSize: 16, color: '#605e5c' } }} />
        </span>
        <Checkbox
          checked
          onChange={(_, v) => {
            if (!v) {
              setTableFilterFields((prev) => prev.filter((f) => f.field !== entry.field));
            }
          }}
          ariaLabel={defaultLabel}
          styles={{ root: { flex: '0 0 auto' } }}
        />
        <Stack tokens={{ childrenGap: 4 }} styles={{ root: { flex: '1 1 200px', minWidth: 0 } }}>
          <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
            <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>{defaultLabel}</Text>
            {o ? (
              <Text variant="small" styles={{ root: { color: '#a19f9d', fontFamily: 'monospace' } }}>{o.meta.MappedType}</Text>
            ) : (
              <Text variant="small" styles={{ root: { color: '#a19f9d', fontFamily: 'monospace' } }}>(campo não listado)</Text>
            )}
          </Stack>
          <TextField
            label="Rótulo do filtro"
            value={entry.label ?? defaultLabel}
            onChange={(_, v) =>
              setTableFilterFields((prev) =>
                prev.map((f) => (f.field === entry.field ? { ...f, label: v ?? '' } : f))
              )
            }
            styles={{ root: { maxWidth: 280 } }}
          />
          {o && (o.meta.MappedType === 'choice' || o.meta.MappedType === 'multichoice') && o.meta.Choices && (
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Opções: {o.meta.Choices.join(' · ')}
            </Text>
          )}
        </Stack>
      </Stack>
    );
  };

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

  const projectRuleFieldOptions: IDropdownOption[] = useMemo(() => {
    const empty: IDropdownOption = { key: '', text: '— selecione o campo —' };
    return [
      empty,
      ...options.map((o) => ({
        key: o.meta.InternalName,
        text: `${o.meta.Title} (${o.meta.InternalName})`,
      })),
    ];
  }, [options]);

  const addProjectColumn = (): void => {
    setProjectColumns((prev) => [
      ...prev,
      {
        id: `col_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
        title: `Coluna ${prev.length + 1}`,
        rules: [],
      },
    ]);
  };

  const updateProjectColumn = (index: number, part: Partial<IProjectManagementColumnConfig>): void => {
    setProjectColumns((prev) => {
      const next = prev.slice();
      if (next[index]) next[index] = { ...next[index], ...part };
      return next;
    });
  };

  const removeProjectColumn = (index: number): void => {
    setProjectColumns((prev) => prev.filter((_, i) => i !== index));
  };

  const addProjectRule = (columnIndex: number): void => {
    setProjectColumns((prev) => {
      const next = prev.slice();
      const current = next[columnIndex];
      if (!current) return prev;
      const rules = current.rules ? current.rules.slice() : [];
      rules.push({
        id: `rule_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
        field: '',
        value: '',
      });
      next[columnIndex] = { ...current, rules };
      return next;
    });
  };

  const updateProjectRule = (columnIndex: number, ruleIndex: number, part: Partial<IProjectManagementRuleConfig>): void => {
    setProjectColumns((prev) => {
      const next = prev.slice();
      const current = next[columnIndex];
      if (!current) return prev;
      const rules = current.rules ? current.rules.slice() : [];
      if (!rules[ruleIndex]) return prev;
      rules[ruleIndex] = { ...rules[ruleIndex], ...part };
      next[columnIndex] = { ...current, rules };
      return next;
    });
  };

  const removeProjectRule = (columnIndex: number, ruleIndex: number): void => {
    setProjectColumns((prev) => {
      const next = prev.slice();
      const current = next[columnIndex];
      if (!current) return prev;
      next[columnIndex] = {
        ...current,
        rules: (current.rules ?? []).filter((_, i) => i !== ruleIndex),
      };
      return next;
    });
  };

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

  const addListRowAction = (): void => {
    setRowActions((prev) => [
      ...prev,
      {
        id: `act_${Date.now()}`,
        title: 'Abrir',
        iconPreset: 'view',
        urlTemplate: '',
        scope: 'icon',
      },
    ]);
  };
  const updateListRowAction = (index: number, part: Partial<IListRowActionConfig>): void => {
    setRowActions((prev) => {
      const next = prev.slice();
      if (next[index]) next[index] = { ...next[index], ...part };
      return next;
    });
  };
  const updateActionVisibility = (index: number, part: Partial<IListRowActionVisibility>): void => {
    setRowActions((prev) => {
      const next = prev.slice();
      if (!next[index]) return prev;
      const cur = next[index].visibility ?? {};
      next[index] = { ...next[index], visibility: { ...cur, ...part } };
      return next;
    });
  };
  const removeListRowAction = (index: number): void => {
    setRowActions((prev) => prev.filter((_, i) => i !== index));
  };

  const appendLayoutCssColor = (property: 'background' | 'color' | 'border-color'): void => {
    const line = `${property}: ${layoutColor};`;
    setLayoutCssText((prev) => {
      const t = prev.trim();
      return t ? `${t}\n${line}` : line;
    });
  };

  const appendCardCssColor = (property: 'background' | 'color' | 'border-color'): void => {
    const line = `${property}: ${cardColor};`;
    setCardCssText((prev) => {
      const t = prev.trim();
      return t ? `${t}\n${line}` : line;
    });
  };

  const appendFilterCssColor = (property: 'background' | 'color' | 'border-color'): void => {
    const line = `${property}: ${filterColor};`;
    setFilterCssText((prev) => {
      const t = prev.trim();
      return t ? `${t}\n${line}` : line;
    });
  };

  const appendViewModeCssColor = (property: 'background' | 'color' | 'border-color'): void => {
    const line = `${property}: ${viewModeColor};`;
    setViewModeCssText((prev) => {
      const t = prev.trim();
      return t ? `${t}\n${line}` : line;
    });
  };
  const appendRuleCssColor = (ruleId: string, index: number, property: 'background' | 'color' | 'border-color'): void => {
    const c = ruleColorMap[ruleId] ?? '#0078d4';
    const line = `${property}: ${c};`;
    setRowStyleRules((prev) => {
      const next = prev.slice();
      const current = next[index];
      if (!current) return prev;
      const body = current.rowCss.trim();
      next[index] = { ...current, rowCss: body ? `${body}\n${line}` : line };
      return next;
    });
  };

  const buildSavePayload = useCallback(() => {
    const columns: IListViewColumnConfig[] = [];
    for (let i = 0; i < options.length; i++) {
      const o = options[i];
      if (!o.selected) continue;
      if (EXPANDABLE.indexOf(o.meta.MappedType) !== -1) {
        const keys =
          o.meta.MappedType === 'lookup' || o.meta.MappedType === 'user'
            ? o.expandFieldsSelected.length > 0
              ? o.expandFieldsSelected
              : [(o.expandField ?? 'Title').trim() || 'Title']
            : [(o.expandField ?? 'Title').trim() || 'Title'];
        const expandOpts = getExpandFieldOptions(o.meta);
        const labelFor = (k: string): string => {
          const hit = expandOpts.find((x) => String(x.key) === k);
          return `${o.meta.Title} – ${hit?.text ?? k}`;
        };
        for (let j = 0; j < keys.length; j++) {
          const ek = (keys[j] ?? 'Title').trim() || 'Title';
          const spanPart = sanitizeColumnSpanMap(o.columnSpanByBreakpoint);
          columns.push({
            field: o.meta.InternalName,
            label: keys.length > 1 ? labelFor(ek) : o.label.trim() ? o.label : o.meta.Title,
            expandField: ek,
            ...(spanPart ? { columnSpanByBreakpoint: spanPart } : {}),
          });
        }
      } else {
        const spanPart = sanitizeColumnSpanMap(o.columnSpanByBreakpoint);
        columns.push({
          field: o.meta.InternalName,
          label: o.label.trim() ? o.label : o.meta.Title,
          ...(spanPart ? { columnSpanByBreakpoint: spanPart } : {}),
        });
      }
    }
    const nextPagination: IPaginationConfig = {
      ...pagination,
      enabled: paginationEnabled,
      pageSize,
      layout: paginationLayout,
      pageSizeOptions: pagination.pageSizeOptions?.length ? pagination.pageSizeOptions : PAGE_SIZE_OPTIONS,
    };
    const cssTrim = layoutCssText.trim();
    const cardCssTrim = cardCssText.trim();
    const filterCssTrim = filterCssText.trim();
    const viewModeCssTrim = viewModeCssText.trim();
    const nextProjectColumns = projectColumns
      .map((col, index) => ({
        id: (col.id || `col_${index + 1}`).trim(),
        title: col.title.trim(),
        rules: (col.rules ?? [])
          .map((rule, ruleIndex) => ({
            id: (rule.id || `rule_${index + 1}_${ruleIndex + 1}`).trim(),
            field: rule.field.trim(),
            value: rule.value,
          }))
          .filter((rule) => rule.field),
      }))
      .filter((col) => col.id && col.title);
    const nextProjectManagement =
      mode === 'projectManagement'
        ? {
            columns: nextProjectColumns,
          }
        : projectManagement;
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
    const nextListRowActions: IListRowActionConfig[] = [];
    for (let i = 0; i < rowActions.length; i++) {
      const a = rowActions[i];
      const title = a.title.trim();
      const urlTemplate = a.urlTemplate.trim();
      const id = (a.id || '').trim() || `act_${Date.now()}_${i}`;
      if (!title || !urlTemplate) continue;
      const iconPreset: TListRowActionIconPreset =
        a.iconPreset === 'view' || a.iconPreset === 'edit' || a.iconPreset === 'link' || a.iconPreset === 'custom'
          ? a.iconPreset
          : 'link';
      const custom =
        iconPreset === 'custom' && (a.customIconName ?? '').trim() ? { customIconName: (a.customIconName ?? '').trim() } : {};
      nextListRowActions.push({
        id,
        title,
        iconPreset,
        ...custom,
        urlTemplate,
        openInNewTab: a.openInNewTab === true,
        scope: a.scope === 'wholeRow' ? 'wholeRow' : 'icon',
        ...(a.visibility ? { visibility: a.visibility } : {}),
      });
    }
    const { listDefaultDisplayMode: _carryListDefault, viewModePicker: _omitVmPicker, ...carryRest } = carryListView;
    const effectiveListCardEnabled =
      mode === 'list' ? listCardViewEnabled : !!(carryListView.listCardViewEnabled ?? false);
    const effectiveListCardDefault: TListViewDisplayMode =
      mode === 'list'
        ? listDefaultDisplayMode
        : carryListView.listDefaultDisplayMode === 'cards'
          ? 'cards'
          : 'table';
    const mappedFilters = tableFilterFields
      .filter((f) => f.field.trim())
      .map((f) => ({
        field: f.field.trim(),
        ...(f.label?.trim() ? { label: f.label.trim() } : {}),
        ...(f.placement === 'advanced' ? { placement: 'advanced' as const } : {}),
      }));
    const nextTableFilterFields: ITableFilterFieldConfig[] = normalizeTableFilterFieldsOrderLocal(mappedFilters);
    const chromeButtonsSanitized = sanitizeListViewChromeButtons(listChromeButtons);
    const nextSort: IListViewSortConfig | null =
      listSortField.trim().length > 0 ? { field: listSortField.trim(), ascending: listSortAscending } : null;
    const listViewOut: IListViewConfig = {
      ...carryRest,
      columns,
      sort: nextSort,
      viewModes,
      activeViewModeId,
      pdfExportEnabled,
      listCardViewEnabled: effectiveListCardEnabled,
      customTableCssSlots: undefined,
      ...(cssTrim ? { customTableCss: cssTrim } : { customTableCss: undefined }),
      ...(cardCssTrim ? { customCardCss: cardCssTrim } : { customCardCss: undefined }),
      ...(filterCssTrim ? { customFilterCss: filterCssTrim } : { customFilterCss: undefined }),
      ...(viewModeCssTrim ? { customViewModeCss: viewModeCssTrim } : { customViewModeCss: undefined }),
      ...(nextRowRules.length > 0 ? { tableRowStyleRules: nextRowRules } : { tableRowStyleRules: undefined }),
      ...(nextListRowActions.length > 0 ? { listRowActions: nextListRowActions } : { listRowActions: undefined }),
      ...(nextTableFilterFields.length > 0 ? { tableFilterFields: nextTableFilterFields } : { tableFilterFields: undefined }),
      ...(tableAdvancedFiltersTitle.trim()
        ? { tableAdvancedFiltersTitle: tableAdvancedFiltersTitle.trim() }
        : { tableAdvancedFiltersTitle: undefined }),
      ...(effectiveListCardEnabled && effectiveListCardDefault === 'cards'
        ? { listDefaultDisplayMode: 'cards' as const }
        : {}),
      ...(listViewModePicker !== 'dropdown' ? { viewModePicker: listViewModePicker } : {}),
      ...(viewModeDefaultRules.length > 0
        ? { viewModeDefaultRules: viewModeDefaultRules.map((r) => ({ ...r })) }
        : { viewModeDefaultRules: undefined }),
      ...(chromeButtonsSanitized?.length ? { chromeButtons: chromeButtonsSanitized } : { chromeButtons: undefined }),
    };
    return {
      listView: listViewOut,
      pagination: nextPagination,
      pdfTemplate: localPdfTemplate,
      projectManagement: nextProjectManagement,
    };
  }, [
    options,
    pagination,
    paginationEnabled,
    pageSize,
    paginationLayout,
    layoutCssText,
    cardCssText,
    filterCssText,
    viewModeCssText,
    projectColumns,
    mode,
    projectManagement,
    rowStyleRules,
    rowActions,
    tableFilterFields,
    tableAdvancedFiltersTitle,
    carryListView,
    viewModes,
    activeViewModeId,
    pdfExportEnabled,
    listCardViewEnabled,
    listDefaultDisplayMode,
    localPdfTemplate,
    listViewModePicker,
    viewModeDefaultRules,
    lookupListFields,
    listChromeButtons,
    listSortField,
    listSortAscending,
  ]);

  const tableJsonPreviewRef = useRef(buildSavePayload());
  tableJsonPreviewRef.current = buildSavePayload();
  useEffect(() => {
    if (jsonOpen) {
      setJsonPanelText(JSON.stringify(tableJsonPreviewRef.current, null, 2));
      setJsonPanelErr(undefined);
    }
  }, [jsonOpen]);

  const applyJsonFromPanel = useCallback(() => {
    setJsonPanelErr(undefined);
    try {
      const parsed: unknown = JSON.parse(jsonPanelText);
      const fallbackPagination: IPaginationConfig = {
        ...pagination,
        enabled: paginationEnabled,
        pageSize,
        layout: paginationLayout,
        pageSizeOptions: pagination.pageSizeOptions?.length ? pagination.pageSizeOptions : PAGE_SIZE_OPTIONS,
      };
      const bundle = sanitizeListTableEditorBundle(
        parsed,
        {
          listView: carryListView,
          pagination: fallbackPagination,
          pdfTemplate: localPdfTemplate,
          projectManagement,
        },
        mode
      );
      if (!bundle) {
        setJsonPanelErr('JSON inválido ou estrutura não reconhecida.');
        return;
      }
      if (options.length === 0 && bundle.listView.columns.length > 0) {
        setJsonPanelErr('Aguarde o carregamento dos campos da lista antes de aplicar colunas.');
        return;
      }
      setCarryListView(bundle.listView);
      setPaginationEnabled(bundle.pagination.enabled);
      setPageSize(bundle.pagination.pageSize);
      setPaginationLayout(bundle.pagination.layout ?? 'buttons');
      setLocalPdfTemplate(bundle.pdfTemplate);
      if (mode === 'projectManagement' && bundle.projectManagement) {
        setProjectColumns(
          bundle.projectManagement.columns?.length ? bundle.projectManagement.columns : DEFAULT_PROJECT_COLUMNS
        );
      }
      setPdfExportEnabled(bundle.listView.pdfExportEnabled ?? false);
      setListCardViewEnabled(bundle.listView.listCardViewEnabled ?? false);
      setListDefaultDisplayMode(bundle.listView.listDefaultDisplayMode === 'cards' ? 'cards' : 'table');
      setLayoutCssText(mergeCustomTableCss(bundle.listView.customTableCssSlots, bundle.listView.customTableCss));
      setCardCssText(bundle.listView.customCardCss ?? '');
      setViewModes(bundle.listView.viewModes?.length ? bundle.listView.viewModes : DEFAULT_VIEW_MODES_FALLBACK);
      setActiveViewModeId(bundle.listView.activeViewModeId ?? 'all');
      setViewModeDefaultRules(bundle.listView.viewModeDefaultRules?.map((r) => ({ ...r })) ?? []);
      setListViewModePicker(normalizeViewModePicker(bundle.listView.viewModePicker));
      setRowStyleRules([...(bundle.listView.tableRowStyleRules ?? [])]);
      setRowActions([...(bundle.listView.listRowActions ?? [])]);
      setTableFilterFields(normalizeTableFilterFieldsOrderLocal(bundle.listView.tableFilterFields?.slice() ?? []));
      setTableAdvancedFiltersTitle(bundle.listView.tableAdvancedFiltersTitle?.trim() ?? '');
      setListChromeButtons((bundle.listView.chromeButtons ?? []).map((b) => ({ ...b })));
      setListSortField(bundle.listView.sort?.field?.trim() ?? '');
      setListSortAscending(bundle.listView.sort?.ascending !== false);
      setOptions((prev) => (prev.length ? applyColumnsToOptions(prev, bundle.listView.columns) : prev));
      setListTabListaSectionOpen({});
      setJsonPanelText(JSON.stringify(bundle, null, 2));
    } catch (e) {
      setJsonPanelErr(e instanceof Error ? e.message : String(e));
    }
  }, [
    jsonPanelText,
    carryListView,
    pagination,
    paginationEnabled,
    pageSize,
    paginationLayout,
    localPdfTemplate,
    projectManagement,
    mode,
    options.length,
  ]);

  const handleSave = (): void => {
    const p = buildSavePayload();
    onSave(p.listView, p.pagination, p.pdfTemplate, p.projectManagement);
    onDismiss();
  };

  const viewModeDefaultOptions: IDropdownOption[] = viewModes.map((m) => ({ key: m.id, text: m.label }));
  const addViewModeDefaultRule = (): void => {
    const firstId = viewModes[0]?.id ?? 'all';
    setViewModeDefaultRules((prev) => [...prev, { viewModeId: firstId }]);
  };
  const updateViewModeDefaultRule = (index: number, patch: Partial<IListViewModeDefaultRule>): void => {
    setViewModeDefaultRules((prev) => {
      const next = prev.slice();
      next[index] = { ...next[index], ...patch };
      return next;
    });
  };
  const removeViewModeDefaultRule = (index: number): void => {
    setViewModeDefaultRules((prev) => prev.filter((_, i) => i !== index));
  };
  const moveViewModeDefaultRule = (index: number, dir: -1 | 1): void => {
    setViewModeDefaultRules((prev) => {
      const j = index + dir;
      if (j < 0 || j >= prev.length) return prev;
      const next = prev.slice();
      const t = next[index];
      next[index] = next[j];
      next[j] = t;
      return next;
    });
  };
  const startViewModeAdd = (): void => {
    setViewModeEditLabel('Novo modo');
    setViewModeEditFilters([]);
    setViewModeEditAccess(undefined);
    setViewModeEditingId(`custom_${Date.now()}`);
  };
  const startViewModeEdit = (m: IListViewModeConfig): void => {
    setViewModeEditingId(m.id);
    setViewModeEditLabel(m.label);
    setViewModeEditFilters(m.filters?.length ? m.filters.slice() : []);
    setViewModeEditAccess(m.access);
  };
  const saveViewModeEdit = (): void => {
    if (viewModeEditingId === null) return;
    const next = viewModes.slice();
    let idx = -1;
    for (let i = 0; i < next.length; i++) { if (next[i].id === viewModeEditingId) { idx = i; break; } }
    const entry: IListViewModeConfig = {
      id: viewModeEditingId,
      label: viewModeEditLabel.trim() || viewModeEditingId,
      filters: viewModeEditFilters,
      ...(viewModeEditAccess !== undefined ? { access: viewModeEditAccess } : {}),
    };
    if (idx >= 0) next[idx] = entry;
    else next.push(entry);
    setViewModes(next);
    setViewModeEditingId(null);
    setViewModeEditAccess(undefined);
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

  const cancelViewModeEdit = (): void => {
    setViewModeEditingId(null);
    setViewModeEditAccess(undefined);
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
    <>
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.custom}
      customWidth="72vw"
      styles={{
        main: {
          width: 'min(72vw, calc(100vw - 16px))',
          maxWidth: 'min(72vw, calc(100vw - 16px))',
          borderRadius: '12px 0 0 12px',
          boxShadow: '-8px 0 40px rgba(0, 0, 0, 0.12)',
        },
        scrollableContent: { overflowX: 'hidden', background: '#f5f4f2' },
        content: { overflowX: 'hidden', minWidth: 0, background: '#f5f4f2' },
        header: {
          background: '#f8f7f6',
          borderBottom: '1px solid #e8e6e4',
          flexShrink: 0,
        },
        footer: { borderTop: '1px solid #e8e6e4', background: '#faf9f8', paddingTop: 12 },
      }}
      headerText={
        mode === 'projectManagement'
          ? 'Editar quadro / cards'
          : mode === 'formManager'
            ? 'Colunas do gestor / lista'
            : 'Editar lista / tabela'
      }
      closeButtonAriaLabel="Fechar"
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="Salvar" onClick={handleSave} />
          <DefaultButton text="Cancelar" onClick={onDismiss} />
        </Stack>
      )}
    >
      <div style={{ padding: '8px 16px 24px', minWidth: 0, maxWidth: '100%', boxSizing: 'border-box' }}>
        {(mode === 'list' || mode === 'projectManagement') && (
          <Stack horizontal horizontalAlign="end" styles={{ root: { marginBottom: 8 } }}>
            <Link onClick={() => setJsonOpen(true)}>JSON (ver / colar)</Link>
          </Stack>
        )}
        {loading ? (
          <Stack horizontalAlign="center" tokens={{ childrenGap: 12 }} style={{ padding: 24 }}>
            <Spinner size={SpinnerSize.medium} />
            <Text variant="small">Carregando campos...</Text>
          </Stack>
        ) : (
          <Stack tokens={{ childrenGap: 18 }} styles={{ root: { minWidth: 0, maxWidth: '100%' } }}>
            <Pivot
              selectedKey={listEditorPivotKey}
              onLinkClick={(item) => {
                const key = item?.props.itemKey;
                if (key) setListEditorPivotKey(key);
              }}
              styles={{ root: { marginBottom: 4 } }}
            >
              <PivotItem itemKey="geral" headerText="Geral">
                <Stack tokens={{ childrenGap: 16 }} styles={{ root: { paddingTop: 12, minWidth: 0, maxWidth: '100%' } }}>
            {mode === 'projectManagement' && (
              <>
                <Stack tokens={{ childrenGap: 8 }}>
                  <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                    Quadro Kanban
                  </Text>
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Cada coluna tem um título e várias regras. Quando o card é arrastado para a coluna, todos os campos definidos nas regras são atualizados no item.
                  </Text>
                  {projectColumns.map((col, index) => (
                    <Stack
                      key={col.id}
                      tokens={{ childrenGap: 8 }}
                      styles={{ root: { padding: 12, border: '1px solid #edebe9', borderRadius: 8, background: '#faf9f8' } }}
                    >
                      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                        <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                          Coluna {index + 1}
                        </Text>
                        <IconButton iconProps={{ iconName: 'Delete' }} title="Remover coluna" onClick={() => removeProjectColumn(index)} />
                      </Stack>
                      <TextField
                        label="Título da coluna"
                        value={col.title}
                        onChange={(_, v) => updateProjectColumn(index, { title: v ?? '' })}
                      />
                      <Text variant="smallPlus" styles={{ root: { fontWeight: 600, marginTop: 4 } }}>
                        Regras
                      </Text>
                      {(col.rules ?? []).map((rule, ruleIndex) => (
                        <Stack
                          key={rule.id}
                          horizontal
                          wrap
                          verticalAlign="end"
                          tokens={{ childrenGap: 8 }}
                          styles={{ root: { padding: 8, borderRadius: 6, background: '#fff', border: '1px solid #edebe9' } }}
                        >
                          <Dropdown
                            label="Campo"
                            selectedKey={rule.field || ''}
                            options={projectRuleFieldOptions}
                            onChange={(_, opt) => updateProjectRule(index, ruleIndex, { field: String(opt?.key ?? '') })}
                            styles={{ root: { flex: '1 1 220px', minWidth: 180 } }}
                          />
                          <TextField
                            label="Valor"
                            value={rule.value}
                            onChange={(_, v) => updateProjectRule(index, ruleIndex, { value: v ?? '' })}
                            styles={{ root: { flex: '1 1 180px', minWidth: 140 } }}
                          />
                          <IconButton
                            iconProps={{ iconName: 'Delete' }}
                            title="Remover regra"
                            onClick={() => removeProjectRule(index, ruleIndex)}
                          />
                        </Stack>
                      ))}
                      <DefaultButton text="Adicionar regra" onClick={() => addProjectRule(index)} />
                    </Stack>
                  ))}
                  <DefaultButton text="Adicionar coluna do quadro" iconProps={{ iconName: 'Add' }} onClick={addProjectColumn} />
                </Stack>

                <Separator />
              </>
            )}
            {mode !== 'projectManagement' && (
              <Stack tokens={{ childrenGap: 12 }} styles={{ root: { width: '100%', minWidth: 0 } }}>
                <ListTabListaCollapse
                  title="Paginação"
                  isOpen={listTabListaSectionOpen.pagination === true}
                  onToggle={() =>
                    setListTabListaSectionOpen((p) => ({
                      ...p,
                      pagination: p.pagination === true ? false : true,
                    }))
                  }
                >
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
                </ListTabListaCollapse>
                <ListTabListaCollapse
                  title="Ordenação inicial da lista"
                  isOpen={listTabListaSectionOpen.tableSort === true}
                  onToggle={() =>
                    setListTabListaSectionOpen((p) => ({
                      ...p,
                      tableSort: p.tableSort === true ? false : true,
                    }))
                  }
                >
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Ordem ao carregar e no OData. Id, Created e Modified estão sempre disponíveis; os outros campos
                    seguem as colunas selecionadas (exc. notas, escolha múltipla e lookups multi-valor).
                  </Text>
                  <Dropdown
                    label="Ordenar por"
                    selectedKey={listSortField || ''}
                    options={[{ key: '', text: '— Nenhuma ordenação inicial —' }, ...listSortFieldOptions]}
                    onChange={(_, opt) => setListSortField(opt?.key ? String(opt.key) : '')}
                    disabled={listSortFieldOptions.length === 0}
                    styles={{ root: { maxWidth: '100%', minWidth: 0 } }}
                  />
                  <ChoiceGroup
                    label="Direção"
                    options={LIST_SORT_DIRECTION_OPTIONS}
                    selectedKey={listSortAscending ? 'asc' : 'desc'}
                    onChange={(_, opt) => opt && setListSortAscending(opt.key === 'asc')}
                    disabled={!listSortField.trim()}
                  />
                </ListTabListaCollapse>
                {mode === 'list' && (
                  <ListTabListaCollapse
                    title="Vista em tabela ou cartões"
                    isOpen={listTabListaSectionOpen.listTableCards === true}
                    onToggle={() =>
                      setListTabListaSectionOpen((p) => ({
                        ...p,
                        listTableCards: p.listTableCards === true ? false : true,
                      }))
                    }
                  >
               
                    <Stack tokens={{ childrenGap: 6 }}>
                      <Checkbox
                        label="Permitir alternar entre tabela e cartões na lista"
                        checked={listCardViewEnabled}
                        onChange={(_, v) => {
                          const on = !!v;
                          setListCardViewEnabled(on);
                          if (!on) setListDefaultDisplayMode('table');
                        }}
                      />
                      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                        Mostra na barra da lista a opção Tabela ou Cartões (as mesmas colunas em grelha de cartões).
                      </Text>
                      {listCardViewEnabled && (
                        <Dropdown
                          label="Visualização ao abrir"
                          selectedKey={listDefaultDisplayMode}
                          options={[
                            { key: 'table', text: 'Tabela' },
                            { key: 'cards', text: 'Cartões' },
                          ]}
                          onChange={(_: React.FormEvent, opt?: IDropdownOption) => {
                            const k = opt?.key as TListViewDisplayMode | undefined;
                            if (k === 'table' || k === 'cards') setListDefaultDisplayMode(k);
                          }}
                          styles={{ root: { maxWidth: 280 } }}
                        />
                      )}
                    </Stack>
                  </ListTabListaCollapse>
                )}
                <ListTabListaCollapse
                  title="Modos de visualização"
                  isOpen={listTabListaSectionOpen.viewModes === true}
                  onToggle={() =>
                    setListTabListaSectionOpen((p) => ({
                      ...p,
                      viewModes: p.viewModes === true ? false : true,
                    }))
                  }
                >
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Ex.: Todas (sem filtro), Minhas (Author/Id eq [Me]), ou filtros customizados. O usuário alterna entre eles na lista.
                  </Text>
                  <ChoiceGroup
                    label="Controle na lista"
                    selectedKey={listViewModePicker}
                    options={VIEW_MODE_PICKER_CHOICES}
                    onChange={(_, opt) => {
                      setListViewModePicker(normalizeViewModePicker(opt?.key));
                    }}
                    styles={{
                      flexContainer: { display: 'flex', flexWrap: 'wrap', columnGap: '12px', rowGap: '4px' },
                    }}
                  />
                  <Dropdown
                    label="Modo padrão"
                    options={viewModeDefaultOptions}
                    selectedKey={activeViewModeId}
                    onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => opt && setActiveViewModeId(String(opt.key))}
                    styles={{ root: { maxWidth: 280 } }}
                  />
                  <Text variant="small" styles={{ root: { fontWeight: 600, marginTop: 14, display: 'block' } }}>
                    Modo inicial por grupo ou utilizador
                  </Text>
                  <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block', marginBottom: 8 } }}>
                    Ordem importa: a primeira regra em que o utilizador se enquadra e o modo lhe é visível define o separador ao abrir. Se nenhuma servir, usa-se &quot;Modo padrão&quot;. Regra sem restrição (sem grupos/pessoas) aplica-se a quem vê esse modo.
                  </Text>
                  {viewModeDefaultRules.map((rule, idx) => {
                    const restrict = rule.access !== undefined;
                    return (
                      <div
                        key={idx}
                        style={{ border: '1px solid #edebe9', borderRadius: 6, padding: 12, marginBottom: 8, background: '#fff' }}
                      >
                        <Stack horizontal verticalAlign="end" wrap tokens={{ childrenGap: 8 }}>
                          <Dropdown
                            label={idx === 0 ? 'Modo' : undefined}
                            selectedKey={rule.viewModeId}
                            options={viewModeDefaultOptions}
                            onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) =>
                              opt && updateViewModeDefaultRule(idx, { viewModeId: String(opt.key) })
                            }
                            styles={{ root: { flex: '1 1 200px', minWidth: 0, maxWidth: 280 } }}
                          />
                          <IconButton
                            iconProps={{ iconName: 'ChevronUp' }}
                            title="Subir"
                            disabled={idx === 0}
                            onClick={() => moveViewModeDefaultRule(idx, -1)}
                          />
                          <IconButton
                            iconProps={{ iconName: 'ChevronDown' }}
                            title="Descer"
                            disabled={idx === viewModeDefaultRules.length - 1}
                            onClick={() => moveViewModeDefaultRule(idx, 1)}
                          />
                          <IconButton
                            iconProps={{ iconName: 'Delete' }}
                            title="Remover regra"
                            onClick={() => removeViewModeDefaultRule(idx)}
                          />
                        </Stack>
                        <Toggle
                          label="Restringir a grupos ou pessoas"
                          checked={restrict}
                          onChange={(_, v) => {
                            if (v) updateViewModeDefaultRule(idx, { access: {} });
                            else updateViewModeDefaultRule(idx, { access: undefined });
                          }}
                          styles={{ root: { marginTop: 8 } }}
                        />
                        {restrict ? (
                          <ViewModeAccessSection
                            value={rule.access}
                            onChange={(next) => updateViewModeDefaultRule(idx, { access: next })}
                            pageWebServerRelativeUrl={pageWebServerRelativeUrl}
                            listWebServerRelativeUrl={listWebServerRelativeUrl}
                          />
                        ) : null}
                      </div>
                    );
                  })}
                  <DefaultButton text="Adicionar regra de modo inicial" onClick={addViewModeDefaultRule} />
                  {viewModes.map((m) => {
                    const accessLine = accessSummary(m.access);
                    return (
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
                          <ViewModeAccessSection
                            value={viewModeEditAccess}
                            onChange={setViewModeEditAccess}
                            pageWebServerRelativeUrl={pageWebServerRelativeUrl}
                            listWebServerRelativeUrl={listWebServerRelativeUrl}
                          />
                          <Stack horizontal tokens={{ childrenGap: 8 }}>
                            <PrimaryButton text="Salvar" onClick={saveViewModeEdit} />
                            <DefaultButton text="Cancelar" onClick={cancelViewModeEdit} />
                          </Stack>
                        </Stack>
                      ) : (
                        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                          <Stack tokens={{ childrenGap: 2 }}>
                            <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>{m.label}</Text>
                            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{viewModeFilterSummary(m.filters)}</Text>
                            {accessLine ? (
                              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{accessLine}</Text>
                            ) : null}
                          </Stack>
                          <Stack horizontal tokens={{ childrenGap: 4 }}>
                            <IconButton iconProps={{ iconName: 'Edit' }} title="Editar" onClick={() => startViewModeEdit(m)} />
                            <IconButton iconProps={{ iconName: 'Delete' }} title="Remover" onClick={() => removeViewMode(m.id)} disabled={m.id === 'all' || m.id === 'mine'} />
                          </Stack>
                        </Stack>
                      )}
                    </div>
                  );
                  })}
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
                        <ViewModeAccessSection
                          value={viewModeEditAccess}
                          onChange={setViewModeEditAccess}
                          pageWebServerRelativeUrl={pageWebServerRelativeUrl}
                          listWebServerRelativeUrl={listWebServerRelativeUrl}
                        />
                        <Stack horizontal tokens={{ childrenGap: 8 }}>
                          <PrimaryButton text="Adicionar modo" onClick={saveViewModeEdit} />
                          <DefaultButton text="Cancelar" onClick={cancelViewModeEdit} />
                        </Stack>
                      </Stack>
                    </div>
                  )}
                  {viewModeEditingId === null && <DefaultButton text="Adicionar modo de visualização" onClick={startViewModeAdd} />}
                </ListTabListaCollapse>
                <ListTabListaCollapse
                  title="Filtros da tabela"
                  isOpen={listTabListaSectionOpen.filterFields === true}
                  onToggle={() =>
                    setListTabListaSectionOpen((p) => ({
                      ...p,
                      filterFields: p.filterFields === true ? false : true,
                    }))
                  }
                >
            
                  <TextField
                    label="Título do agrupamento de filtros avançados"
                    description="Texto do controlo que expande ou recolhe os filtros avançados na barra da lista."
                    placeholder="Filtros avançados"
                    value={tableAdvancedFiltersTitle}
                    onChange={(_, v) => setTableAdvancedFiltersTitle(v ?? '')}
                    styles={{ root: { maxWidth: 400 } }}
                  />
                  {options.length === 0 && <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>Carregando campos…</Text>}
                  <ListTabListaCollapse
                    title="Filtros fixos (sempre visíveis na lista)"
                    isOpen={filterEditorFixedOpen}
                    onToggle={() => setFilterEditorFixedOpen((v) => !v)}
                    zoneDrop={{
                      highlight: tableFilterZoneDropHighlight === 'fixed',
                      onDragOver: onTableFilterZoneDragOver('fixed'),
                      onDrop: onTableFilterZoneDrop('fixed'),
                    }}
                  >
                    {tableFiltersPartitioned.fixed.length === 0 ? (
                      <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>Nenhum filtro fixo. Largue aqui um filtro arrastado ou adicione nos campos abaixo.</Text>
                    ) : (
                      tableFiltersPartitioned.fixed.map((entry, localIdx) =>
                        renderConfiguredTableFilterRow(entry, 'fixed', localIdx)
                      )
                    )}
                  </ListTabListaCollapse>
                  <ListTabListaCollapse
                    title="Filtros avançados (no collapse da lista)"
                    isOpen={filterEditorAdvancedOpen}
                    onToggle={() => setFilterEditorAdvancedOpen((v) => !v)}
                    zoneDrop={{
                      highlight: tableFilterZoneDropHighlight === 'advanced',
                      onDragOver: onTableFilterZoneDragOver('advanced'),
                      onDrop: onTableFilterZoneDrop('advanced'),
                    }}
                  >
                    {tableFiltersPartitioned.advanced.length === 0 ? (
                      <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>Nenhum filtro avançado. Largue aqui um filtro arrastado para o colocar neste grupo.</Text>
                    ) : (
                      tableFiltersPartitioned.advanced.map((entry, localIdx) =>
                        renderConfiguredTableFilterRow(entry, 'advanced', localIdx)
                      )
                    )}
                  </ListTabListaCollapse>
                  <Separator styles={{ root: { marginTop: 4, marginBottom: 4 } }} />
                  <Text variant="small" styles={{ root: { fontWeight: 600, color: '#605e5c' } }}>Outros campos da lista</Text>
                  {filterFieldUnselectedOptions.map((o) => {
                    const fieldKey = getTableFilterFieldKey(o);
                    const defaultLabel = o.meta.Title;
                    return (
                      <Stack
                        key={o.meta.InternalName}
                        horizontal
                        wrap
                        tokens={{ childrenGap: 12 }}
                        verticalAlign="center"
                        styles={{ root: { padding: '8px 0', borderBottom: '1px solid #f3f2f1', width: '100%', minWidth: 0 } }}
                      >
                        <span style={{ width: 22, flexShrink: 0 }} aria-hidden />
                        <Checkbox
                          checked={false}
                          onChange={(_, v) => {
                            if (v) {
                              setTableFilterFields((prev) => {
                                const { fixed, advanced } = partitionTableFilterFields(prev);
                                return mergeTableFilterPartitions([...fixed, { field: fieldKey, label: defaultLabel }], advanced);
                              });
                            }
                          }}
                          ariaLabel={o.meta.Title}
                          styles={{ root: { flex: '0 0 auto' } }}
                        />
                        <Stack tokens={{ childrenGap: 4 }} styles={{ root: { flex: '1 1 200px', minWidth: 0 } }}>
                          <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
                            <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>{o.meta.Title}</Text>
                            <Text variant="small" styles={{ root: { color: '#a19f9d', fontFamily: 'monospace' } }}>{o.meta.MappedType}</Text>
                          </Stack>
                          {(o.meta.MappedType === 'choice' || o.meta.MappedType === 'multichoice') && o.meta.Choices && (
                            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                              Opções: {o.meta.Choices.join(' · ')}
                            </Text>
                          )}
                        </Stack>
                      </Stack>
                    );
                  })}
                </ListTabListaCollapse>
                <ListTabListaCollapse
                  title="Botões na barra da tabela"
                  isOpen={listTabListaSectionOpen.chromeButtons === true}
                  onToggle={() =>
                    setListTabListaSectionOpen((p) => ({
                      ...p,
                      chromeButtons: p.chromeButtons === true ? false : true,
                    }))
                  }
                >
                 
                  {listChromeButtons.map((b, idx) => (
                    <Stack
                      key={`${b.id}_${idx}`}
                      tokens={{ childrenGap: 8 }}
                      styles={{
                        root: {
                          padding: 12,
                          border: '1px solid #edebe9',
                          borderRadius: 8,
                          background: '#faf9f8',
                          width: '100%',
                          minWidth: 0,
                          boxSizing: 'border-box',
                        },
                      }}
                    >
                      <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 8 }}>
                        <TextField
                          label="Id"
                          value={b.id}
                          onChange={(_, v) => {
                            const id = v ?? '';
                            setListChromeButtons((prev) => prev.map((x, i) => (i === idx ? { ...x, id } : x)));
                          }}
                          styles={{ root: { flex: '1 1 140px', minWidth: 120, maxWidth: 220 } }}
                        />
                        <TextField
                          label="Texto"
                          value={b.label}
                          onChange={(_, v) =>
                            setListChromeButtons((prev) => prev.map((x, i) => (i === idx ? { ...x, label: v ?? '' } : x)))
                          }
                          styles={{ root: { flex: '2 1 180px', minWidth: 140 } }}
                        />
                        <Dropdown
                          label="Posição"
                          selectedKey={b.slot}
                          options={LIST_VIEW_CHROME_SLOT_OPTIONS}
                          onChange={(_, opt) => {
                            const key = opt?.key as TListViewChromeButtonSlot | undefined;
                            if (!key) return;
                            setListChromeButtons((prev) => prev.map((x, i) => (i === idx ? { ...x, slot: key } : x)));
                          }}
                          styles={{ root: { flex: '1 1 260px', minWidth: 220 } }}
                        />
                        <Dropdown
                          label="Ação"
                          selectedKey={b.actionKind}
                          options={LIST_CHROME_ACTION_OPTIONS}
                          onChange={(_, opt) => {
                            const ak = opt?.key === 'reload' ? 'reload' : 'redirect';
                            setListChromeButtons((prev) =>
                              prev.map((x, i) =>
                                i === idx
                                  ? {
                                      ...x,
                                      actionKind: ak,
                                      ...(ak === 'reload' ? { url: undefined, openInNewTab: undefined } : {}),
                                    }
                                  : x
                              )
                            );
                          }}
                          styles={{ root: { flex: '0 0 auto', width: 180 } }}
                        />
                        <TextField
                          label="Ordem"
                          value={b.order !== undefined ? String(b.order) : ''}
                          onChange={(_, v) => {
                            const t = (v ?? '').trim();
                            setListChromeButtons((prev) =>
                              prev.map((x, i) =>
                                i === idx
                                  ? {
                                      ...x,
                                      ...(t === '' ? { order: undefined } : { order: Number.parseInt(t, 10) || 0 }),
                                    }
                                  : x
                              )
                            );
                          }}
                          styles={{ root: { width: 72 } }}
                        />
                        <IconButton
                          iconProps={{ iconName: 'Delete' }}
                          title="Remover"
                          onClick={() => setListChromeButtons((prev) => prev.filter((_, i) => i !== idx))}
                        />
                      </Stack>
                      {b.actionKind === 'redirect' ? (
                        <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="center">
                          <TextField
                            label="URL"
                            value={b.url ?? ''}
                            onChange={(_, v) =>
                              setListChromeButtons((prev) => prev.map((x, i) => (i === idx ? { ...x, url: v ?? '' } : x)))
                            }
                            styles={{ root: { flex: '1 1 280px', minWidth: 200 } }}
                          />
                          <Checkbox
                            label="Nova aba"
                            checked={b.openInNewTab === true}
                            onChange={(_, v) =>
                              setListChromeButtons((prev) =>
                                prev.map((x, i) =>
                                  i === idx ? { ...x, openInNewTab: v === true ? true : undefined } : x
                                )
                              )
                            }
                          />
                        </Stack>
                      ) : null}
                      <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                        <Dropdown
                          label="Variante"
                          selectedKey={b.variant === 'primary' ? 'primary' : 'default'}
                          options={[
                            { key: 'default', text: 'Padrão' },
                            { key: 'primary', text: 'Primário' },
                          ]}
                          onChange={(_, opt) => {
                            const vr = opt?.key === 'primary' ? 'primary' : 'default';
                            setListChromeButtons((prev) => prev.map((x, i) => (i === idx ? { ...x, variant: vr } : x)));
                          }}
                          styles={{ root: { width: 140 } }}
                        />
                        <TextField
                          label="Ícone Fluent (opcional)"
                          value={b.iconName ?? ''}
                          onChange={(_, v) =>
                            setListChromeButtons((prev) =>
                              prev.map((x, i) =>
                                i === idx
                                  ? {
                                      ...x,
                                      ...(v?.trim() ? { iconName: v.trim() } : { iconName: undefined }),
                                    }
                                  : x
                              )
                            )
                          }
                          styles={{ root: { flex: '1 1 160px', minWidth: 120 } }}
                        />
                        <TextField
                          label="CSS extra (opcional)"
                          value={b.css ?? ''}
                          multiline
                          rows={2}
                          onChange={(_, v) =>
                            setListChromeButtons((prev) =>
                              prev.map((x, i) =>
                                i === idx ? { ...x, ...(v?.trim() ? { css: v } : { css: undefined }) } : x
                              )
                            )
                          }
                          styles={{ root: { flex: '2 1 240px', minWidth: 160 } }}
                        />
                      </Stack>
                    </Stack>
                  ))}
                  <DefaultButton
                    text="Adicionar botão"
                    iconProps={{ iconName: 'Add' }}
                    onClick={() => setListChromeButtons((prev) => [...prev, createDefaultListChromeButton()])}
                  />
                </ListTabListaCollapse>
                <ListTabListaCollapse
                  title="Colunas da tabela"
                  isOpen={listTabListaSectionOpen.columns === true}
                  onToggle={() =>
                    setListTabListaSectionOpen((p) => ({
                      ...p,
                      columns: p.columns === true ? false : true,
                    }))
                  }
                >
                
                  {options.map((o, rowIndex) => (
                    <Stack
                      key={o.meta.InternalName}
                      horizontal
                      wrap
                      tokens={{ childrenGap: 12 }}
                      verticalAlign="start"
                      onDragOver={(e) => onListTableColumnDragOver(e, rowIndex)}
                      onDragLeave={onListTableColumnDragLeave}
                      onDrop={onListTableColumnDrop(rowIndex)}
                      onDragEnd={onListTableColumnDragLeave}
                      styles={{
                        root: {
                          padding: '8px 0',
                          borderBottom: '1px solid #f3f2f1',
                          width: '100%',
                          minWidth: 0,
                          boxSizing: 'border-box',
                          borderTop:
                            listColumnDragOverIndex === rowIndex ? '2px solid #0078d4' : undefined,
                          backgroundColor:
                            listColumnDragOverIndex === rowIndex ? 'rgba(0, 120, 212, 0.04)' : undefined,
                        },
                      }}
                    >
                      <span
                        draggable={o.selected}
                        onDragStart={onListTableColumnDragStart(rowIndex, o.selected)}
                        title={o.selected ? 'Arrastar para reordenar' : undefined}
                        style={{
                          cursor: o.selected ? 'grab' : 'default',
                          display: 'flex',
                          alignItems: 'flex-start',
                          paddingTop: 6,
                          opacity: o.selected ? 1 : 0.35,
                          flexShrink: 0,
                        }}
                      >
                        <Icon iconName="GripperBarVertical" styles={{ root: { fontSize: 16, color: '#605e5c' } }} />
                      </span>
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
                        {EXPANDABLE.indexOf(o.meta.MappedType) !== -1 &&
                          o.selected &&
                          (o.meta.MappedType === 'lookup' || o.meta.MappedType === 'user') && (
                          <Stack
                            tokens={{ childrenGap: 6 }}
                            styles={{
                              root: {
                                marginTop: 6,
                                paddingLeft: 20,
                                borderLeft: '3px solid #edebe9',
                              },
                            }}
                          >
                            <Text variant="small" styles={{ root: { fontWeight: 600, color: '#605e5c' } }}>
                              Campos da lista ligada
                            </Text>
                            {getExpandFieldOptions(o.meta).map((opt) => {
                              const k = String(opt.key);
                              return (
                                <Checkbox
                                  key={k}
                                  label={opt.text}
                                  checked={o.expandFieldsSelected.indexOf(k) !== -1}
                                  onChange={(_, v) => toggleLookupExpandField(o.meta.InternalName, k, !!v)}
                                />
                              );
                            })}
                          </Stack>
                        )}
                        {EXPANDABLE.indexOf(o.meta.MappedType) !== -1 &&
                          o.selected &&
                          (o.meta.MappedType === 'lookupmulti' || o.meta.MappedType === 'usermulti') && (
                          <Dropdown
                            label="Campo expandido (lookup multi / utilizadores)"
                            selectedKey={o.expandField || 'Title'}
                            options={getExpandFieldOptions(o.meta)}
                            onChange={(_, opt) => setExpandField(o.meta.InternalName, (opt?.key as string) ?? 'Title')}
                            styles={{ root: { maxWidth: '100%' } }}
                          />
                        )}
                        {mode === 'list' && o.selected ? (
                          <Stack tokens={{ childrenGap: 6 }} styles={{ root: { marginTop: 10, maxWidth: '100%' } }}>
                            <Text variant="small" styles={{ root: { fontWeight: 600, color: '#605e5c' } }}>
                              Largura por ecrã (grelha 12; padrão 12)
                            </Text>
                            <Pivot>
                              {LIST_VIEW_COLUMN_BREAKPOINT_ORDER.map((bp) => {
                                const idx = LIST_VIEW_COLUMN_BREAKPOINT_ORDER.indexOf(bp);
                                const inherited =
                                  idx === 0
                                    ? 12
                                    : resolveListColumnSpanAtBreakpoint(idx - 1, o.columnSpanByBreakpoint);
                                const resolved = resolveListColumnSpanAtBreakpoint(idx, o.columnSpanByBreakpoint);
                                return (
                                  <PivotItem
                                    key={bp}
                                    headerText={LIST_VIEW_COLUMN_BREAKPOINT_LABEL[bp]}
                                    headerButtonProps={{
                                      'aria-label': LIST_VIEW_COLUMN_BREAKPOINT_LABEL[bp],
                                      title: `≥ ${LIST_VIEW_COLUMN_BREAKPOINT_MIN_PX[bp]}px`,
                                    }}
                                  >
                                    <Dropdown
                                      label="Em colunas de 12"
                                      selectedKey={resolved}
                                      options={LIST_COLUMN_SPAN_OPTIONS}
                                      onChange={(_, opt) => {
                                        const v = clampListColumnSpan(Number(opt?.key) || 12);
                                        setOptions((prev) =>
                                          prev.map((row) => {
                                            if (row.meta.InternalName !== o.meta.InternalName) return row;
                                            const map = { ...row.columnSpanByBreakpoint };
                                            if (v === inherited) delete map[bp];
                                            else map[bp] = v;
                                            return { ...row, columnSpanByBreakpoint: map };
                                          })
                                        );
                                      }}
                                      styles={{ root: { maxWidth: 130 } }}
                                    />
                                  </PivotItem>
                                );
                              })}
                            </Pivot>
                          </Stack>
                        ) : null}
                      </Stack>
                      <Text variant="small" styles={{ root: { color: '#a19f9d', flex: '0 0 auto', maxWidth: '100%', wordBreak: 'break-word' } }}>
                        {o.meta.MappedType}
                      </Text>
                    </Stack>
                  ))}
                </ListTabListaCollapse>
              </Stack>
            )}

            {mode === 'projectManagement' && (
              <>
                <Stack tokens={{ childrenGap: 8 }}>
                  <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                    Campos do card
                  </Text>
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Marque os campos que deseja mostrar dentro do card. O primeiro campo selecionado vira o título do card.
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
                      {EXPANDABLE.indexOf(o.meta.MappedType) !== -1 &&
                        o.selected &&
                        (o.meta.MappedType === 'lookup' || o.meta.MappedType === 'user') && (
                        <Stack
                          tokens={{ childrenGap: 6 }}
                          styles={{
                            root: {
                              marginTop: 6,
                              paddingLeft: 20,
                              borderLeft: '3px solid #edebe9',
                            },
                          }}
                        >
                          <Text variant="small" styles={{ root: { fontWeight: 600, color: '#605e5c' } }}>
                            Campos da lista ligada
                          </Text>
                          {getExpandFieldOptions(o.meta).map((opt) => {
                            const k = String(opt.key);
                            return (
                              <Checkbox
                                key={k}
                                label={opt.text}
                                checked={o.expandFieldsSelected.indexOf(k) !== -1}
                                onChange={(_, v) => toggleLookupExpandField(o.meta.InternalName, k, !!v)}
                              />
                            );
                          })}
                        </Stack>
                      )}
                      {EXPANDABLE.indexOf(o.meta.MappedType) !== -1 &&
                        o.selected &&
                        (o.meta.MappedType === 'lookupmulti' || o.meta.MappedType === 'usermulti') && (
                        <Dropdown
                          label="Campo expandido (lookup multi / utilizadores)"
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
              </>
            )}
              {mode !== 'projectManagement' ? (
                <ListTabListaCollapse
                  title="Ações na linha"
                  isOpen={listTabListaSectionOpen.rowActions === true}
                  onToggle={() =>
                    setListTabListaSectionOpen((p) => ({
                      ...p,
                      rowActions: p.rowActions === true ? false : true,
                    }))
                  }
                >
                  <Stack tokens={{ childrenGap: 14 }} styles={{ root: { minWidth: 0, maxWidth: '100%' } }}>
                    <Stack
                      tokens={{ childrenGap: 10 }}
                      styles={{ root: { width: '100%', flexShrink: 0, minHeight: 0 } }}
                    >
           
                      <Text
                        variant="small"
                        block
                        styles={{
                          root: {
                            display: 'block',
                            width: '100%',
                            margin: 0,
                            color: '#605e5c',
                            lineHeight: 1.55,
                            whiteSpace: 'normal',
                            wordBreak: 'break-word',
                            overflowWrap: 'break-word',
                          },
                        }}
                      >
                        {`Campos: {{ID}}, {{ Title }} (duplas chaves), {Id}, {Title}; lookup {Autor/Title}; tokens [me], [siteurl], [query:chave], etc.`}
                      </Text>
                    </Stack>
                    {rowActions.map((act, ai) => (
                      <Stack
                        key={act.id}
                        tokens={{ childrenGap: 10 }}
                        styles={{
                          root: {
                            padding: 14,
                            background: '#faf9f8',
                            borderRadius: 8,
                            border: '1px solid #edebe9',
                            maxWidth: '100%',
                          },
                        }}
                      >
                        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                          <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                            Ação {ai + 1}
                          </Text>
                          <IconButton iconProps={{ iconName: 'Delete' }} title="Remover" ariaLabel="Remover ação" onClick={() => removeListRowAction(ai)} />
                        </Stack>
                        <TextField label="Título (tooltip)" value={act.title} onChange={(_, v) => updateListRowAction(ai, { title: v ?? '' })} styles={{ root: { maxWidth: '100%' } }} />
                        <Dropdown
                          label="Ícone"
                          selectedKey={act.iconPreset}
                          options={LIST_ROW_ICON_PRESET_OPTIONS}
                          onChange={(_, opt) =>
                            opt && updateListRowAction(ai, { iconPreset: String(opt.key) as TListRowActionIconPreset })
                          }
                          styles={{ root: { maxWidth: 320 } }}
                        />
                        {act.iconPreset === 'custom' && (
                          <TextField
                            label="Nome do ícone Fluent"
                            value={act.customIconName ?? ''}
                            onChange={(_, v) => updateListRowAction(ai, { customIconName: v ?? '' })}
                            placeholder="Ex.: Mail, Share"
                            styles={{ root: { maxWidth: '100%' } }}
                          />
                        )}
                        <TextField
                          label="URL"
                          multiline
                          resizable
                          rows={3}
                          value={act.urlTemplate}
                          onChange={(_, v) => updateListRowAction(ai, { urlTemplate: v ?? '' })}
                          placeholder="https://...?Form={{ID}} ou ?id={Id}"
                          styles={{ root: { maxWidth: '100%' } }}
                        />
                        <Checkbox
                          label="Abrir em nova aba"
                          checked={act.openInNewTab === true}
                          onChange={(_, v) => updateListRowAction(ai, { openInNewTab: !!v })}
                        />
                        <ChoiceGroup
                          label="Clique"
                          selectedKey={act.scope}
                          options={LIST_ROW_SCOPE_OPTIONS}
                          onChange={(_, opt) =>
                            opt && updateListRowAction(ai, { scope: opt.key as IListRowActionConfig['scope'] })
                          }
                        />

                        {/* Seção de Visibilidade */}
                        <ListTabListaCollapse
                          title={`Visibilidade${(act.visibility?.allowedGroupIds?.length ?? 0) + (act.visibility?.allowedUserLogins?.length ?? 0) + (act.visibility?.fieldRules?.length ?? 0) > 0 ? ' ✓' : ''}`}
                          isOpen={visibilitySectionOpen[act.id] === true}
                          onToggle={() => setVisibilitySectionOpen((p) => ({ ...p, [act.id]: !p[act.id] }))}
                        >
                          <Text variant="small" styles={{ root: { color: '#605e5c', lineHeight: 1.5 } }}>
                            Sem configuração = visível para todos. Com configuração, o botão só aparece se <strong>identidade</strong> (grupo OU usuário) e <strong>regras de campo</strong> (AND) passarem.
                          </Text>

                          {/* Grupos */}
                          <Text variant="smallPlus" styles={{ root: { fontWeight: 600, marginTop: 4 } }}>Grupos (IDs numéricos SharePoint)</Text>
                          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                            Usuário em qualquer um dos grupos → visível. Deixe vazio para sem restrição de grupo.
                          </Text>
                          {(act.visibility?.allowedGroupIds ?? []).map((gid, gi) => (
                            <Stack key={gi} horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                              <Stack.Item grow={1}>
                                <TextField
                                  value={gid}
                                  placeholder="Ex.: 5"
                                  onChange={(_, v) => {
                                    const ids = [...(act.visibility?.allowedGroupIds ?? [])];
                                    ids[gi] = v ?? '';
                                    updateActionVisibility(ai, { allowedGroupIds: ids });
                                  }}
                                  styles={{ root: { maxWidth: '100%' } }}
                                />
                              </Stack.Item>
                              <IconButton
                                iconProps={{ iconName: 'Delete' }}
                                title="Remover"
                                onClick={() => {
                                  const ids = (act.visibility?.allowedGroupIds ?? []).filter((_, i) => i !== gi);
                                  updateActionVisibility(ai, { allowedGroupIds: ids.length ? ids : undefined });
                                }}
                              />
                            </Stack>
                          ))}
                          <DefaultButton
                            text="Adicionar grupo"
                            iconProps={{ iconName: 'Add' }}
                            styles={{ root: { alignSelf: 'flex-start' } }}
                            onClick={() => updateActionVisibility(ai, { allowedGroupIds: [...(act.visibility?.allowedGroupIds ?? []), ''] })}
                          />

                          {/* Usuários específicos */}
                          <Text variant="smallPlus" styles={{ root: { fontWeight: 600, marginTop: 8 } }}>Usuários (loginName)</Text>
                          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                            Ex.: i:0#.f|membership|joao@empresa.com
                          </Text>
                          {(act.visibility?.allowedUserLogins ?? []).map((login, li) => (
                            <Stack key={li} horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                              <Stack.Item grow={1}>
                                <TextField
                                  value={login}
                                  placeholder="i:0#.f|membership|..."
                                  onChange={(_, v) => {
                                    const logins = [...(act.visibility?.allowedUserLogins ?? [])];
                                    logins[li] = v ?? '';
                                    updateActionVisibility(ai, { allowedUserLogins: logins });
                                  }}
                                  styles={{ root: { maxWidth: '100%' } }}
                                />
                              </Stack.Item>
                              <IconButton
                                iconProps={{ iconName: 'Delete' }}
                                title="Remover"
                                onClick={() => {
                                  const logins = (act.visibility?.allowedUserLogins ?? []).filter((_, i) => i !== li);
                                  updateActionVisibility(ai, { allowedUserLogins: logins.length ? logins : undefined });
                                }}
                              />
                            </Stack>
                          ))}
                          <DefaultButton
                            text="Adicionar usuário"
                            iconProps={{ iconName: 'Add' }}
                            styles={{ root: { alignSelf: 'flex-start' } }}
                            onClick={() => updateActionVisibility(ai, { allowedUserLogins: [...(act.visibility?.allowedUserLogins ?? []), ''] })}
                          />

                          {/* Regras de campo */}
                          <Text variant="smallPlus" styles={{ root: { fontWeight: 600, marginTop: 8 } }}>Regras de campo (AND)</Text>
                          <Text variant="small" styles={{ root: { color: '#605e5c', lineHeight: 1.5 } }}>
                            Tokens de valor: <span style={{ fontFamily: 'monospace' }}>[Me.Id]</span> = ID do usuário logado, <span style={{ fontFamily: 'monospace' }}>[Me.Login]</span> = loginName. Ex.: campo <em>Author/Id</em> igual a <em>[Me.Id]</em>.
                          </Text>
                          {(act.visibility?.fieldRules ?? []).map((rule, ri) => (
                            <Stack key={ri} tokens={{ childrenGap: 6 }} styles={{ root: { padding: 8, background: '#f3f9ff', borderRadius: 6, border: '1px solid #c7e0f4' } }}>
                              <Stack horizontal tokens={{ childrenGap: 8 }} wrap verticalAlign="end">
                                <Stack.Item styles={{ root: { flex: '1 1 140px', minWidth: 120 } }}>
                                  <TextField
                                    label="Campo"
                                    value={rule.field}
                                    placeholder="Author/Id"
                                    onChange={(_, v) => {
                                      const rules = [...(act.visibility?.fieldRules ?? [])];
                                      rules[ri] = { ...rules[ri], field: v ?? '' };
                                      updateActionVisibility(ai, { fieldRules: rules });
                                    }}
                                  />
                                </Stack.Item>
                                <Stack.Item styles={{ root: { flex: '1 1 120px', minWidth: 100 } }}>
                                  <Dropdown
                                    label="Operador"
                                    selectedKey={rule.op}
                                    options={ROW_ACTION_FIELD_RULE_OP_OPTIONS}
                                    onChange={(_, opt) => {
                                      if (!opt) return;
                                      const rules = [...(act.visibility?.fieldRules ?? [])];
                                      rules[ri] = { ...rules[ri], op: opt.key as TListRowActionFieldRuleOp };
                                      updateActionVisibility(ai, { fieldRules: rules });
                                    }}
                                  />
                                </Stack.Item>
                                <Stack.Item styles={{ root: { flex: '1 1 160px', minWidth: 120 } }}>
                                  <TextField
                                    label="Valor"
                                    value={rule.value}
                                    placeholder="[Me.Id]"
                                    onChange={(_, v) => {
                                      const rules = [...(act.visibility?.fieldRules ?? [])];
                                      rules[ri] = { ...rules[ri], value: v ?? '' };
                                      updateActionVisibility(ai, { fieldRules: rules });
                                    }}
                                  />
                                </Stack.Item>
                                <IconButton
                                  iconProps={{ iconName: 'Delete' }}
                                  title="Remover regra"
                                  onClick={() => {
                                    const rules = (act.visibility?.fieldRules ?? []).filter((_, i) => i !== ri);
                                    updateActionVisibility(ai, { fieldRules: rules.length ? rules : undefined });
                                  }}
                                />
                              </Stack>
                            </Stack>
                          ))}
                          <DefaultButton
                            text="Adicionar regra de campo"
                            iconProps={{ iconName: 'Add' }}
                            styles={{ root: { alignSelf: 'flex-start' } }}
                            onClick={() => {
                              const newRule: IListRowActionFieldRule = { field: '', op: 'eq', value: '[Me.Id]' };
                              updateActionVisibility(ai, { fieldRules: [...(act.visibility?.fieldRules ?? []), newRule] });
                            }}
                          />
                        </ListTabListaCollapse>
                      </Stack>
                    ))}
                    <DefaultButton text="Adicionar ação" iconProps={{ iconName: 'Add' }} onClick={addListRowAction} />
                  </Stack>
                </ListTabListaCollapse>
              ) : null}
              {showPdfExcelConfigTabs ? (
                <ListTabListaCollapse
                  title="Exportar PDF e Excel"
                  isOpen={listTabListaSectionOpen.exportPdfExcel === true}
                  onToggle={() =>
                    setListTabListaSectionOpen((p) => ({
                      ...p,
                      exportPdfExcel: p.exportPdfExcel === true ? false : true,
                    }))
                  }
                >
                  <Stack tokens={{ childrenGap: 16 }} styles={{ root: { minWidth: 0, maxWidth: '100%' } }}>
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
                    <Separator />
                    <Text variant="medium" styles={{ root: { color: '#605e5c', fontWeight: 600 } }}>
                      Excel
                    </Text>
                    <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                      Exportação para Excel em breve.
                    </Text>
                  </Stack>
                </ListTabListaCollapse>
              ) : null}
                </Stack>
              </PivotItem>
              <PivotItem itemKey="css" headerText="CSS">
                <Stack tokens={{ childrenGap: 10 }} styles={{ root: { paddingTop: 12, paddingBottom: 8, minWidth: 0, maxWidth: '100%' } }}>
                  <ListTabListaCollapse
                    title="CSS da tabela"
                    isOpen={layoutSectionOpen.tableCss === true}
                    onToggle={() => setLayoutSectionOpen((p) => ({ ...p, tableCss: !p.tableCss }))}
                  >
                    <Text variant="small" styles={{ root: { color: '#323130', lineHeight: 1.55 } }}>
                      Vazio = estilo padrão moderno da tabela. Com texto, só o seu CSS é aplicado. A pré-visualização inclui as regras de linha.
                    </Text>
                    <Stack horizontal wrap verticalAlign="start" tokens={{ childrenGap: 16 }}>
                      <Stack styles={{ root: { flex: '1 1 480px', minWidth: 320 } }} tokens={{ childrenGap: 8 }}>
                        <TextField
                          label="CSS da tabela"
                          multiline
                          resizable
                          rows={18}
                          value={layoutCssText}
                          onChange={(_, v) => setLayoutCssText(v ?? '')}
                          placeholder={
                            `.${DINAMIC_SX_TABLE_CLASS.headerCell} { background: #f3f2f1; }\n` +
                            `.${DINAMIC_SX_TABLE_CLASS.row}:nth-child(even) { background: #faf9f8; }\n` +
                            `.${DINAMIC_SX_TABLE_CLASS.cell}[data-field="Title"] { font-weight: 600; }`
                          }
                          styles={{ root: { maxWidth: '100%' } }}
                        />
                        <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 8 }}>
                          <input
                            type="color"
                            value={normalizeHexColor(layoutColor, '#0078d4')}
                            onChange={(e) => setLayoutColor(e.target.value)}
                            aria-label="Cor para CSS da tabela"
                            style={{ width: 40, height: 32, border: '1px solid #edebe9', borderRadius: 4, background: '#fff', cursor: 'pointer' }}
                          />
                          <TextField
                            label="Cor (hex)"
                            value={layoutColor}
                            onChange={(_, v) => setLayoutColor((v ?? '').trim() || '#000000')}
                            styles={{ root: { width: 130 } }}
                          />
                          <DefaultButton text="Inserir background" onClick={() => appendLayoutCssColor('background')} />
                          <DefaultButton text="Inserir color" onClick={() => appendLayoutCssColor('color')} />
                          <DefaultButton text="Inserir border-color" onClick={() => appendLayoutCssColor('border-color')} />
                        </Stack>
                        <Stack
                          tokens={{ childrenGap: 4 }}
                          styles={{ root: { padding: 10, border: '1px solid #edebe9', borderRadius: 6, background: '#faf9f8', maxHeight: 220, overflowY: 'auto' } }}
                        >
                          {TABLE_LAYOUT_EDITOR_ROWS.map((row) => (
                            <Text key={row.slot} variant="small" styles={{ root: { color: '#605e5c' } }}>
                              <span style={{ fontFamily: 'monospace', color: '#0078d4' }}>.{DINAMIC_SX_TABLE_CLASS[row.slot]}</span> — {row.title}
                            </Text>
                          ))}
                        </Stack>
                      </Stack>
                      <Stack styles={{ root: { flex: '1 1 360px', minWidth: 280, maxWidth: '100%' } }}>
                        <TableLayoutLivePreview cssText={layoutPreviewCss} rulePreviewTokens={layoutPreviewRuleTokens} />
                      </Stack>
                    </Stack>
                  </ListTabListaCollapse>

                  {/* ── Regras de linha ── */}
                  <ListTabListaCollapse
                    title="Regras de linha"
                    isOpen={layoutSectionOpen.rowRules === true}
                    onToggle={() => setLayoutSectionOpen((p) => ({ ...p, rowRules: !p.rowRules }))}
                  >
                    <Text variant="small" styles={{ root: { color: '#605e5c', lineHeight: 1.55 } }}>
                      Quando o valor de um campo atender à condição, o CSS é aplicado em todas as células da linha (marcador <span style={{ fontFamily: 'monospace' }}>data-dinamic-rules</span>).
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
                            <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 8 }}>
                              <input
                                type="color"
                                value={normalizeHexColor(ruleColorMap[rule.id], '#0078d4')}
                                onChange={(e) =>
                                  setRuleColorMap((prev) => ({
                                    ...prev,
                                    [rule.id]: e.target.value,
                                  }))
                                }
                                aria-label={`Selecionar cor para regra ${ri + 1}`}
                                style={{ width: 40, height: 32, border: '1px solid #edebe9', borderRadius: 4, background: '#fff', cursor: 'pointer' }}
                              />
                              <TextField
                                label="Cor (hex)"
                                value={ruleColorMap[rule.id] ?? '#0078d4'}
                                onChange={(_, v) =>
                                  setRuleColorMap((prev) => ({
                                    ...prev,
                                    [rule.id]: (v ?? '').trim() || '#000000',
                                  }))
                                }
                                styles={{ root: { width: 130 } }}
                              />
                              <DefaultButton text="Inserir background" onClick={() => appendRuleCssColor(rule.id, ri, 'background')} />
                              <DefaultButton text="Inserir color" onClick={() => appendRuleCssColor(rule.id, ri, 'color')} />
                              <DefaultButton text="Inserir border-color" onClick={() => appendRuleCssColor(rule.id, ri, 'border-color')} />
                            </Stack>
                            {/* <Text variant="small" styles={{ root: { fontFamily: 'monospace', color: '#605e5c' } }}>
                              Marcador: data-dinamic-rules~=&quot;{toTableRowRuleDataToken(rule.id)}&quot;
                            </Text> */}
                          </Stack>
                        );
                      })}
                      <DefaultButton text="Adicionar regra" iconProps={{ iconName: 'Add' }} onClick={addRowStyleRule} />
                  </ListTabListaCollapse>

                  {/* ── CSS Cards ── */}
                  <ListTabListaCollapse
                    title="CSS dos cards"
                    isOpen={layoutSectionOpen.cardCss === true}
                    onToggle={() => setLayoutSectionOpen((p) => ({ ...p, cardCss: !p.cardCss }))}
                  >
                    <Text variant="small" styles={{ root: { color: '#323130', lineHeight: 1.55 } }}>
                      Personaliza a visualização em cards (ativa quando o usuário alterna para Cards). Use os seletores abaixo.
                    </Text>
                    <Stack tokens={{ childrenGap: 8 }}>
                      <TextField
                        label="CSS dos cards"
                        multiline
                        resizable
                        rows={14}
                        value={cardCssText}
                        onChange={(_, v) => setCardCssText(v ?? '')}
                        placeholder={
                          `.dinamicSxCard { border-radius: 12px; }\n` +
                          `.dinamicSxCardTitle { color: #0078d4; font-size: 15px; }\n` +
                          `.dinamicSxCardLabel { color: #605e5c; }\n` +
                          `.dinamicSxCardGrid { gap: 16px; }`
                        }
                        styles={{ root: { maxWidth: '100%' } }}
                      />
                      <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 8 }}>
                        <input
                          type="color"
                          value={normalizeHexColor(cardColor, '#0078d4')}
                          onChange={(e) => setCardColor(e.target.value)}
                          aria-label="Cor para CSS dos cards"
                          style={{ width: 40, height: 32, border: '1px solid #edebe9', borderRadius: 4, background: '#fff', cursor: 'pointer' }}
                        />
                        <TextField
                          label="Cor (hex)"
                          value={cardColor}
                          onChange={(_, v) => setCardColor((v ?? '').trim() || '#000000')}
                          styles={{ root: { width: 130 } }}
                        />
                        <DefaultButton text="Inserir background" onClick={() => appendCardCssColor('background')} />
                        <DefaultButton text="Inserir color" onClick={() => appendCardCssColor('color')} />
                        <DefaultButton text="Inserir border-color" onClick={() => appendCardCssColor('border-color')} />
                      </Stack>
                      <Stack
                        tokens={{ childrenGap: 4 }}
                        styles={{ root: { padding: 10, border: '1px solid #edebe9', borderRadius: 6, background: '#faf9f8' } }}
                      >
                        {([
                          { cls: 'dinamicSxCardGrid', desc: 'Container da grade de cards' },
                          { cls: 'dinamicSxCard', desc: 'Card individual' },
                          { cls: 'dinamicSxCardTitle', desc: 'Primeiro campo (título)' },
                          { cls: 'dinamicSxCardField', desc: 'Linha de campo (label + valor)' },
                          { cls: 'dinamicSxCardLabel', desc: 'Rótulo do campo' },
                          { cls: 'dinamicSxCardValue', desc: 'Valor do campo' },
                          { cls: 'dinamicSxCardActions', desc: 'Área de botões de ação' },
                        ] as const).map((r) => (
                          <Text key={r.cls} variant="small" styles={{ root: { color: '#605e5c' } }}>
                            <span style={{ fontFamily: 'monospace', color: '#0078d4' }}>.{r.cls}</span> — {r.desc}
                          </Text>
                        ))}
                      </Stack>
                    </Stack>
                  </ListTabListaCollapse>

                  {/* ── CSS Filtros ── */}
                  <ListTabListaCollapse
                    title="CSS dos filtros"
                    isOpen={layoutSectionOpen.filterCss === true}
                    onToggle={() => setLayoutSectionOpen((p) => ({ ...p, filterCss: !p.filterCss }))}
                  >
                    <Text variant="small" styles={{ root: { color: '#323130', lineHeight: 1.55 } }}>
                      Personaliza a barra de filtros. Use os seletores abaixo para estilizar o container e os controles individuais.
                    </Text>
                    <Stack tokens={{ childrenGap: 8 }}>
                      <TextField
                        label="CSS dos filtros"
                        multiline
                        resizable
                        rows={12}
                        value={filterCssText}
                        onChange={(_, v) => setFilterCssText(v ?? '')}
                        placeholder={
                          `.dinamicSxFilterBar { background: #f3f2f1; border-radius: 4px; }\n` +
                          `.dinamicSxFilterControl label { color: #0078d4; font-weight: 700; }\n` +
                          `.dinamicSxFilterControl input { border-color: #0078d4; }`
                        }
                        styles={{ root: { maxWidth: '100%' } }}
                      />
                      <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 8 }}>
                        <input
                          type="color"
                          value={normalizeHexColor(filterColor, '#0078d4')}
                          onChange={(e) => setFilterColor(e.target.value)}
                          aria-label="Cor para CSS dos filtros"
                          style={{ width: 40, height: 32, border: '1px solid #edebe9', borderRadius: 4, background: '#fff', cursor: 'pointer' }}
                        />
                        <TextField
                          label="Cor (hex)"
                          value={filterColor}
                          onChange={(_, v) => setFilterColor((v ?? '').trim() || '#000000')}
                          styles={{ root: { width: 130 } }}
                        />
                        <DefaultButton text="Inserir background" onClick={() => appendFilterCssColor('background')} />
                        <DefaultButton text="Inserir color" onClick={() => appendFilterCssColor('color')} />
                        <DefaultButton text="Inserir border-color" onClick={() => appendFilterCssColor('border-color')} />
                      </Stack>
                      <Stack
                        tokens={{ childrenGap: 4 }}
                        styles={{ root: { padding: 10, border: '1px solid #edebe9', borderRadius: 6, background: '#faf9f8' } }}
                      >
                        {([
                          { cls: 'dinamicSxFilterBar', desc: 'Container da barra de filtros' },
                          { cls: 'dinamicSxFilterControl', desc: 'Wrapper de cada controle (label + input)' },
                          { cls: 'dinamicSxFilterControl label', desc: 'Label do controle' },
                          { cls: 'dinamicSxFilterControl input', desc: 'Input de texto / data' },
                          { cls: 'dinamicSxFilterControl .ms-Dropdown', desc: 'Dropdown (choice / boolean)' },
                        ] as const).map((r) => (
                          <Text key={r.cls} variant="small" styles={{ root: { color: '#605e5c' } }}>
                            <span style={{ fontFamily: 'monospace', color: '#0078d4' }}>.{r.cls}</span> — {r.desc}
                          </Text>
                        ))}
                      </Stack>
                    </Stack>
                  </ListTabListaCollapse>

                  {/* ── CSS Modos de Visualização ── */}
                  <ListTabListaCollapse
                    title="CSS dos modos de visualização"
                    isOpen={layoutSectionOpen.viewModeCss === true}
                    onToggle={() => setLayoutSectionOpen((p) => ({ ...p, viewModeCss: !p.viewModeCss }))}
                  >
                    <Text variant="small" styles={{ root: { color: '#323130', lineHeight: 1.55 } }}>
                      Estiliza a barra de modos de visualização (abas e dropdown). Use os seletores abaixo.
                    </Text>
                    <Stack tokens={{ childrenGap: 8 }}>
                      <TextField
                        label="CSS dos modos de visualização"
                        multiline
                        resizable
                        rows={12}
                        value={viewModeCssText}
                        onChange={(_, v) => setViewModeCssText(v ?? '')}
                        placeholder={
                          `.dinamicSxViewModeBar { background: #f3f2f1; padding: 8px; border-radius: 4px; }\n` +
                          `.dinamicSxViewModeTab { border-radius: 4px; }\n` +
                          `.dinamicSxViewModeTab[aria-selected="true"] { background: #0078d4; color: #fff; }`
                        }
                        styles={{ root: { maxWidth: '100%' } }}
                      />
                      <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 8 }}>
                        <input
                          type="color"
                          value={normalizeHexColor(viewModeColor, '#0078d4')}
                          onChange={(e) => setViewModeColor(e.target.value)}
                          aria-label="Cor para CSS dos modos de visualização"
                          style={{ width: 40, height: 32, border: '1px solid #edebe9', borderRadius: 4, background: '#fff', cursor: 'pointer' }}
                        />
                        <TextField
                          label="Cor (hex)"
                          value={viewModeColor}
                          onChange={(_, v) => setViewModeColor((v ?? '').trim() || '#000000')}
                          styles={{ root: { width: 130 } }}
                        />
                        <DefaultButton text="Inserir background" onClick={() => appendViewModeCssColor('background')} />
                        <DefaultButton text="Inserir color" onClick={() => appendViewModeCssColor('color')} />
                        <DefaultButton text="Inserir border-color" onClick={() => appendViewModeCssColor('border-color')} />
                      </Stack>
                      <Stack
                        tokens={{ childrenGap: 4 }}
                        styles={{ root: { padding: 10, border: '1px solid #edebe9', borderRadius: 6, background: '#faf9f8' } }}
                      >
                        {([
                          { cls: 'dinamicSxViewModeBar', desc: 'Container da barra de modos (abas ou dropdown)' },
                          { cls: 'dinamicSxViewModeTab', desc: 'Cada botão de aba' },
                          { cls: 'dinamicSxViewModeTab[aria-selected="true"]', desc: 'Aba ativa' },
                          { cls: 'dinamicSxViewModeDropdown', desc: 'Dropdown de modos de visualização' },
                          { cls: 'dinamicSxViewModeDropdown .ms-Dropdown', desc: 'Controle interno do dropdown' },
                        ] as const).map((r) => (
                          <Text key={r.cls} variant="small" styles={{ root: { color: '#605e5c' } }}>
                            <span style={{ fontFamily: 'monospace', color: '#0078d4' }}>.{r.cls}</span> — {r.desc}
                          </Text>
                        ))}
                      </Stack>
                    </Stack>
                  </ListTabListaCollapse>
                </Stack>
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
    <Panel
      isOpen={jsonOpen}
      type={PanelType.medium}
      headerText="Lista / tabela (JSON)"
      onDismiss={() => setJsonOpen(false)}
    >
      <Text variant="small" styles={{ root: { color: '#605e5c', marginBottom: 8 } }}>
        Objeto com «listView», «pagination» e opcionalmente «pdfTemplate» e «projectManagement». Aplicar carrega no
        painel; Salvar grava na vista.
      </Text>
      {jsonPanelErr && (
        <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 8 } }}>
          {jsonPanelErr}
        </MessageBar>
      )}
      <TextField
        multiline
        rows={22}
        value={jsonPanelText}
        onChange={(_, v) => setJsonPanelText(v ?? '')}
        styles={{ root: { fontFamily: 'monospace', fontSize: 12 } }}
      />
      <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 12 } }}>
        <PrimaryButton text="Aplicar JSON" onClick={() => applyJsonFromPanel()} />
        <DefaultButton text="Fechar" onClick={() => setJsonOpen(false)} />
      </Stack>
    </Panel>
    </>
  );
};
