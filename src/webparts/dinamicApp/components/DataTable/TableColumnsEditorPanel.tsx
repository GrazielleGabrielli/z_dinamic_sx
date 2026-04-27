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
  Pivot,
  PivotItem,
  MessageBar,
  MessageBarType,
} from '@fluentui/react';
import { FieldsService } from '../../../../services';
import type { IFieldMetadata } from '../../../../services';
import type {
  IProjectManagementColumnConfig,
  IProjectManagementConfig,
  IProjectManagementRuleConfig,
  IListViewConfig,
  IListViewColumnConfig,
  IListViewModeConfig,
  IListViewFilterConfig,
  IPaginationConfig,
  IPdfTemplateConfig,
  IListRowActionConfig,
  ITableRowStyleRule,
  TListRowActionIconPreset,
  TTableRowRuleOperator,
  TPaginationLayout,
  TFilterOperator,
  TViewMode,
  TListViewDisplayMode,
} from '../../core/config/types';
import { PdfTemplateEditor } from './PdfTemplateEditor';
import {
  DINAMIC_SX_TABLE_CLASS,
  TABLE_LAYOUT_EDITOR_ROWS,
  mergeCustomTableCss,
  mergeRowStyleRulesCss,
} from './tableLayoutClasses';
import { isNoteFieldMeta } from '../../core/listView';
import { toTableRowRuleDataToken } from '../../core/table/utils/tableRowStyleRuleEval';
import { TableLayoutLivePreview } from './TableLayoutLivePreview';
import { sanitizeListTableEditorBundle } from '../../core/config/validators';

interface ITableColumnsEditorPanelProps {
  isOpen: boolean;
  mode: TViewMode;
  listTitle: string;
  listWebServerRelativeUrl?: string;
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

function applyColumnsToOptions(opts: IFieldOption[], cols: IListViewColumnConfig[]): IFieldOption[] {
  const map = new Map(opts.map((o) => [o.meta.InternalName, o]));
  const ordered: IFieldOption[] = [];
  for (let i = 0; i < cols.length; i++) {
    const c = cols[i];
    const o = map.get(c.field);
    if (!o) continue;
    ordered.push({
      ...o,
      selected: true,
      label: c.label && c.label.trim() ? c.label : o.meta.Title,
      expandField: c.expandField ?? o.expandField,
    });
    map.delete(c.field);
  }
  map.forEach((o) => ordered.push({ ...o, selected: false }));
  return ordered;
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

const DEFAULT_PROJECT_COLUMNS: IProjectManagementColumnConfig[] = [];

function normalizeHexColor(input: string | undefined, fallback: string): string {
  const raw = (input ?? '').trim();
  return /^#([0-9a-fA-F]{6})$/.test(raw) ? raw : fallback;
}

function viewModeFilterSummary(filters: IListViewFilterConfig[]): string {
  if (!filters || filters.length === 0) return 'Sem filtros';
  return filters.map((f) => `${f.field} ${f.operator} "${f.value}"`).join(' e ');
}

type TListTabListaSection = 'pagination' | 'viewModes' | 'columns';

function ListTabListaCollapse(props: {
  title: string;
  isOpen: boolean;
  onToggle: () => void;
  children: React.ReactNode;
}): JSX.Element {
  return (
    <Stack
      styles={{
        root: {
          border: '1px solid #edebe9',
          borderRadius: 10,
          background: '#ffffff',
          boxShadow: '0 1px 2px rgba(0, 0, 0, 0.04)',
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
        styles={{
          root: {
            padding: '10px 12px',
            background: props.isOpen ? '#faf9f8' : '#ffffff',
            borderBottom: props.isOpen ? '1px solid #edebe9' : undefined,
            userSelect: 'none',
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
          style={{
            padding: '14px 14px 16px 18px',
            maxWidth: '100%',
            minWidth: 0,
            width: '100%',
            boxSizing: 'border-box',
            display: 'flex',
            flexDirection: 'column',
            gap: 12,
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
  listView,
  pagination,
  projectManagement,
  pdfTemplate,
  onSave,
  onDismiss,
}) => {
  const lw = listWebServerRelativeUrl?.trim() || undefined;
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
  const [listCardViewEnabled, setListCardViewEnabled] = useState(listView.listCardViewEnabled ?? false);
  const [listDefaultDisplayMode, setListDefaultDisplayMode] = useState<TListViewDisplayMode>(
    listView.listDefaultDisplayMode === 'cards' ? 'cards' : 'table'
  );
  const [layoutCssText, setLayoutCssText] = useState<string>(
    mergeCustomTableCss(listView.customTableCssSlots, listView.customTableCss)
  );
  const [layoutColor, setLayoutColor] = useState<string>('#0078d4');
  const [projectColumns, setProjectColumns] = useState<IProjectManagementColumnConfig[]>(
    projectManagement?.columns?.length ? projectManagement.columns : DEFAULT_PROJECT_COLUMNS
  );
  const [rowStyleRules, setRowStyleRules] = useState<ITableRowStyleRule[]>(() => [
    ...(listView.tableRowStyleRules ?? []),
  ]);
  const [rowActions, setRowActions] = useState<IListRowActionConfig[]>(() => [...(listView.listRowActions ?? [])]);
  const layoutPreviewCss = useMemo(() => {
    const layout = layoutCssText.trim();
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
  const [viewModeEditingId, setViewModeEditingId] = useState<string | null>(null);
  const [viewModeEditLabel, setViewModeEditLabel] = useState('');
  const [viewModeEditFilters, setViewModeEditFilters] = useState<IListViewFilterConfig[]>([]);
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
    setLocalPdfTemplate(pdfTemplate);
    setPdfExportEnabled(listView.pdfExportEnabled ?? false);
    setListCardViewEnabled(listView.listCardViewEnabled ?? false);
    setListDefaultDisplayMode(listView.listDefaultDisplayMode === 'cards' ? 'cards' : 'table');
    setLayoutCssText(mergeCustomTableCss(listView.customTableCssSlots, listView.customTableCss));
    setProjectColumns(projectManagement?.columns?.length ? projectManagement.columns : DEFAULT_PROJECT_COLUMNS);
    setRowStyleRules([...(listView.tableRowStyleRules ?? [])]);
    setRowActions([...(listView.listRowActions ?? [])]);
    setRuleColorMap({});
    setLayoutSubTab('geral');
    setListTabListaSectionOpen({});
  }, [isOpen, listView, pagination, pdfTemplate, projectManagement]);

  const showPdfExcelConfigTabs = mode !== 'list';

  useEffect(() => {
    if (mode !== 'list') return;
    setActiveTab((tab) => (tab === 'pdf' || tab === 'excel' ? 'lista' : tab));
  }, [mode]);

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
    const cssTrim = layoutCssText.trim();
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
      });
    }
    const { listDefaultDisplayMode: _carryListDefault, ...carryRest } = carryListView;
    const listViewOut: IListViewConfig = {
      ...carryRest,
      columns,
      viewModes,
      activeViewModeId,
      pdfExportEnabled,
      listCardViewEnabled,
      customTableCssSlots: undefined,
      ...(cssTrim ? { customTableCss: cssTrim } : { customTableCss: undefined }),
      ...(nextRowRules.length > 0 ? { tableRowStyleRules: nextRowRules } : { tableRowStyleRules: undefined }),
      ...(nextListRowActions.length > 0 ? { listRowActions: nextListRowActions } : { listRowActions: undefined }),
      ...(listCardViewEnabled && listDefaultDisplayMode === 'cards' ? { listDefaultDisplayMode: 'cards' as const } : {}),
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
    projectColumns,
    mode,
    projectManagement,
    rowStyleRules,
    rowActions,
    carryListView,
    viewModes,
    activeViewModeId,
    pdfExportEnabled,
    listCardViewEnabled,
    listDefaultDisplayMode,
    localPdfTemplate,
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
      setViewModes(bundle.listView.viewModes?.length ? bundle.listView.viewModes : DEFAULT_VIEW_MODES_FALLBACK);
      setActiveViewModeId(bundle.listView.activeViewModeId ?? 'all');
      setRowStyleRules([...(bundle.listView.tableRowStyleRules ?? [])]);
      setRowActions([...(bundle.listView.listRowActions ?? [])]);
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
    <>
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.custom}
      customWidth="68vw"
      styles={{
        main: { width: 'min(68vw, calc(100vw - 16px))', maxWidth: 'min(68vw, calc(100vw - 16px))' },
        scrollableContent: { overflowX: 'hidden' },
        content: { overflowX: 'hidden', minWidth: 0 },
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
      <div style={{ paddingTop: 16, minWidth: 0, maxWidth: '100%', boxSizing: 'border-box' }}>
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
          <Stack tokens={{ childrenGap: 16 }} styles={{ root: { minWidth: 0, maxWidth: '100%' } }}>
            <Pivot
              selectedKey={activeTab}
              onLinkClick={(item) => item?.props?.itemKey !== undefined && item?.props?.itemKey !== null && setActiveTab(String(item.props.itemKey))}
              styles={{ root: { marginBottom: 8, flexWrap: 'wrap', maxWidth: '100%' } }}
            >
              <PivotItem itemKey="lista" headerText="Lista">
                <Stack tokens={{ childrenGap: 16 }} styles={{ root: { paddingTop: 8, minWidth: 0, maxWidth: '100%' } }}>
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
                  <Stack tokens={{ childrenGap: 6 }}>
                    <Checkbox
                      label="Permitir visualização em cards na lista"
                      checked={listCardViewEnabled}
                      onChange={(_, v) => {
                        const on = !!v;
                        setListCardViewEnabled(on);
                        if (!on) setListDefaultDisplayMode('table');
                      }}
                    />
                    <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                      Exibe na lista a opção Tabela ou Cards (mesmas colunas em grade de cartões).
                    </Text>
                    {listCardViewEnabled && (
                      <Dropdown
                        label="Visualização inicial"
                        selectedKey={listDefaultDisplayMode}
                        options={[
                          { key: 'table', text: 'Tabela' },
                          { key: 'cards', text: 'Cards' },
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
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Marque as colunas que deseja exibir. Para lookups e usuários, escolha o campo de exibição.
                  </Text>
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
              </>
            )}
                </Stack>
              </PivotItem>
              {mode !== 'projectManagement' ? (
                <PivotItem itemKey="acoes" headerText="Ações">
                  <Stack tokens={{ childrenGap: 14 }} styles={{ root: { paddingTop: 8, paddingBottom: 24, minWidth: 0, maxWidth: '100%' } }}>
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
                        {
                          'Botões por item na tabela e nos cards. Opcionalmente a linha ou o card inteiro abre a URL da ação marcada como "Linha ou card inteiro" (usa a primeira ação assim na lista).'
                        }
                      </Text>
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
                      </Stack>
                    ))}
                    <DefaultButton text="Adicionar ação" iconProps={{ iconName: 'Add' }} onClick={addListRowAction} />
                  </Stack>
                </PivotItem>
              ) : null}
              {showPdfExcelConfigTabs ? (
                <>
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
                </>
              ) : null}
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
                    <Stack tokens={{ childrenGap: 14 }} styles={{ root: { paddingTop: 8, paddingBottom: 24, minWidth: 0, maxWidth: '100%' } }}>
                      <Text variant="small" styles={{ root: { color: '#323130', lineHeight: 1.55 } }}>
                        Use uma única caixa de CSS com seletores das classes da tabela. A pré-visualização ao lado reage ao digitar
                        e inclui o CSS das regras de linha (aba <strong>Regras</strong>).
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
                          <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 4 } }}>
                            <input
                              type="color"
                              value={normalizeHexColor(layoutColor, '#0078d4')}
                              onChange={(e) => setLayoutColor(e.target.value)}
                              aria-label="Selecionar cor para CSS da tabela"
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
                                <span style={{ fontFamily: 'monospace', color: '#0078d4' }}>.{DINAMIC_SX_TABLE_CLASS[row.slot]}</span> - {row.title}
                              </Text>
                            ))}
                          </Stack>
                        </Stack>
                        <Stack styles={{ root: { flex: '1 1 360px', minWidth: 280, maxWidth: '100%' } }}>
                          <TableLayoutLivePreview cssText={layoutPreviewCss} rulePreviewTokens={layoutPreviewRuleTokens} />
                        </Stack>
                      </Stack>
                    </Stack>
                  </PivotItem>
                  <PivotItem itemKey="regras" headerText="Regras">
                    <Stack tokens={{ childrenGap: 14 }} styles={{ root: { paddingTop: 8, paddingBottom: 24, minWidth: 0, maxWidth: '100%' } }}>
                      <Text variant="small" styles={{ root: { color: '#605e5c', lineHeight: 1.55 } }}>
                        Quando o valor exibido de um campo atender à condição, o mesmo CSS é aplicado em <strong>todas as células da linha</strong>{' '}
                        (marcador <span style={{ fontFamily: 'monospace' }}>data-dinamic-rules</span> em cada <span style={{ fontFamily: 'monospace' }}>&lt;td&gt;</span>, para fundo e bordas ficarem corretos com{' '}
                        <span style={{ fontFamily: 'monospace' }}>border-collapse</span>). A comparação usa o mesmo texto da célula (incluindo lookups).
                        Várias regras podem valer ao mesmo tempo; cada uma acrescenta um token no marcador.
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
