import * as React from 'react';
import { useState, useEffect, useMemo, useRef } from 'react';
import {
  Stack,
  Dropdown,
  IDropdownOption,
  ActionButton,
  DefaultButton,
  PrimaryButton,
  TextField,
} from '@fluentui/react';
import {
  IDynamicViewConfig,
  IListPageButtonItemConfig,
  IListViewChromeButtonConfig,
  IListViewConfig,
  IListViewFilterConfig,
  IListViewModeConfig,
  TListViewChromeButtonSlot,
} from '../../core/config/types';
import { TableEngine } from '../../core/table/services/TableEngine';
import type { ITableConfig, ISortConfig } from '../../core/table/types';
import { buildListFilter, buildTableTopFiltersOData, getActiveViewModeFilters } from '../../core/listView';
import {
  filterViewModesForCurrentUser,
  pickFallbackViewModeId,
} from '../../core/listView/viewModeAccess';
import { useViewModeMembership } from '../../core/listView/useViewModeMembership';
import { buildDynamicContext, parseQueryString } from '../../core/dynamicTokens';
import { generateAndDownloadPdf } from '../../core/pdf';
import { ItemsService, UsersService, FieldsService, SYSTEM_METADATA_FIELDS } from '../../../../services';
import { readListItemId } from '../../../../services/items/listItemId';
import { DataTable } from './DataTable';
import { ListItemsCardGrid } from './ListItemsCardGrid';
import { TableCardsLayoutToggle } from './TableCardsLayoutToggle';
import { useResponsiveListTableColumns } from './useResponsiveListTableColumns';
import {
  DINAMIC_SX_TABLE_CLASS,
  mergeRowStyleRulesCss,
  resolveTableLayoutCss,
  scopeCardCssByInstance,
  TABLE_COLUMN_FILTER_PORTAL_CSS,
} from './tableLayoutClasses';
import { ViewModePickerBar } from './ViewModePickerBar';
import { resolveViewModeCss } from './viewModePickerLayouts';
import { DINAMIC_SX_FILTER_CLASS, resolveFilterBarCss } from './filterBarLayouts';
import { DINAMIC_SX_TOOLBAR_CLASS, resolveListToolbarCss } from './listToolbarLayouts';
import { columnODataPath } from '../../core/table/utils/columnODataPath';
import {
  isSafeListRowNavigationUrl,
  resolveListRowActionUrl,
} from '../../core/table/utils/resolveListRowActionUrl';
import type { IDynamicContext } from '../../core/dynamicTokens/types';

const EMPTY_VIEW_MODES: IListViewModeConfig[] = [];

function navigateListChromeButton(
  it: IListPageButtonItemConfig,
  dynamicContext: IDynamicContext,
  rowContext: Record<string, unknown>
): void {
  if (it.actionKind === 'reload') {
    window.location.reload();
    return;
  }
  const template = (it.url ?? '').trim();
  if (!template) return;
  const u = resolveListRowActionUrl(template, rowContext, dynamicContext).trim();
  if (!u || !isSafeListRowNavigationUrl(u)) return;
  if (it.openInNewTab === true) {
    window.open(u, '_blank', 'noopener,noreferrer');
  } else {
    window.location.assign(u);
  }
}

function parseListChromeCss(css: string | undefined): React.CSSProperties | undefined {
  if (!css?.trim()) return undefined;
  const style: Record<string, string> = {};
  css.split(';').forEach((decl) => {
    const idx = decl.indexOf(':');
    if (idx < 0) return;
    const prop = decl.slice(0, idx).trim();
    const val = decl.slice(idx + 1).trim();
    if (!prop || !val) return;
    const camel = prop.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());
    style[camel] = val;
  });
  return Object.keys(style).length > 0 ? (style as React.CSSProperties) : undefined;
}

function sortChromeForSlot(items: IListViewChromeButtonConfig[]): IListViewChromeButtonConfig[] {
  return [...items].sort((a, b) => {
    const oa = a.order ?? 0;
    const ob = b.order ?? 0;
    if (oa !== ob) return oa - ob;
    return a.label.localeCompare(b.label);
  });
}

function listViewToTableConfig(listView: IDynamicViewConfig['listView']): Partial<ITableConfig> {
  const rawCols = listView.columns ?? [];
  const countByField = new Map<string, number>();
  for (let i = 0; i < rawCols.length; i++) {
    const f = rawCols[i].field;
    countByField.set(f, (countByField.get(f) ?? 0) + 1);
  }
  const columns = rawCols.map((c, idx) => {
    const expand = c.expandField?.trim();
    const idSafe = expand ? `${c.field}__${expand}`.replace(/[^\w-]/g, '_') : c.field;
    const dupField = (countByField.get(c.field) ?? 0) > 1;
    const firstIdx = rawCols.findIndex((x) => x.field === c.field);
    const spanMap = c.columnSpanByBreakpoint;
    return {
      id: idSafe,
      internalName: c.field,
      label: c.label ?? c.field,
      visible: true,
      sortable: dupField ? idx === firstIdx : true,
      expandConfig: expand ? { displayField: expand } : undefined,
      ...(spanMap && Object.keys(spanMap).length > 0 ? { columnSpanByBreakpoint: spanMap } : {}),
    };
  });
  return {
    enabled: true,
    columns: columns as ITableConfig['columns'],
    sortable: true,
    defaultSort: listView.sort?.field
      ? { field: listView.sort.field, direction: listView.sort.ascending ? 'asc' : 'desc' }
      : undefined,
    emptyMessage: 'Nenhum item encontrado.',
  };
}

export interface ITableViewProps {
  config: IDynamicViewConfig;
  /** Filtros OData do item do dashboard (card/série) clicado; combinados com modo de visualização e filtros de coluna. */
  dashboardListFilters?: IListViewFilterConfig[];
  instanceScopeId: string;
  /** Site da página (para avaliar grupos ao restringir modos). */
  pageWebServerRelativeUrl?: string;
  /** Notifica alteração do modo de visualização (sincronizar com dashboard vinculado). */
  onActiveViewModeChange?: (viewModeId: string) => void;
  /** Incrementar para limpar todos os filtros internos (coluna + barra de filtros). */
  clearFiltersSignal?: number;
  /** Limpar filtros do dashboard (seleção de card/série). */
  onClearFilters?: () => void;
}

function scopeTableCssByInstance(css: string, scopeClass: string): string {
  return scopeCssSelectorsByInstance(css, scopeClass);
}

function scopeCssSelectorsByInstance(css: string, scopeClass: string): string {
  if (!css.trim()) return '';
  const scope = `.${scopeClass}`;
  return css.replace(/(^|})\s*([^{}]+)\{/g, (match, prefix, selectorPart) => {
    const trimmed = selectorPart.trim();
    if (!trimmed || trimmed.startsWith('@') || trimmed.includes(scope)) {
      return match;
    }
    const scoped = trimmed
      .split(',')
      .map((sel: string) => `${scope} ${sel.trim()}`)
      .join(', ');
    return `${prefix} ${scoped}{`;
  });
}

function scopeFilterCssByInstance(css: string, scopeClass: string): string {
  return scopeCssSelectorsByInstance(css, scopeClass);
}

function scopeViewModeCssByInstance(css: string, scopeClass: string): string {
  if (!css.trim()) return '';
  return css.replace(/\.dinamicSxViewMode/g, `.${scopeClass} .dinamicSxViewMode`);
}

function scopeToolbarCssByInstance(css: string, scopeClass: string): string {
  return scopeCssSelectorsByInstance(css, scopeClass);
}

export const TableView: React.FC<ITableViewProps> = ({
  config,
  dashboardListFilters,
  instanceScopeId,
  pageWebServerRelativeUrl,
  onActiveViewModeChange,
  clearFiltersSignal,
  onClearFilters,
}) => {
  const { dataSource, pagination, listView, tableConfig: tableConfigRaw } = config;
  const listTitle = dataSource.title;
  const listWeb = dataSource.webServerRelativeUrl?.trim() || undefined;

  const tableConfigFromList = useMemo(() => listViewToTableConfig(listView), [listView]);
  const initialTableConfig = useMemo(
    () => (tableConfigRaw && tableConfigRaw.columns?.length ? tableConfigRaw : tableConfigFromList) as Partial<ITableConfig>,
    [tableConfigRaw, tableConfigFromList]
  );

  const [tableConfig, setTableConfig] = useState<ITableConfig | null>(null);
  const [sortConfig, setSortConfig] = useState<ISortConfig | null>(
    () => initialTableConfig.defaultSort ?? null
  );
  const [items, setItems] = useState<Record<string, unknown>[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | undefined>(undefined);
  const [paging, setPaging] = useState<{
    pageIndex: number;
    forwardPivots: number[];
    resetKey: string | null;
  }>({ pageIndex: 0, forwardPivots: [], resetKey: null });
  const [hasNext, setHasNext] = useState(false);
  const [columnFilters, setColumnFilters] = useState<Record<string, string>>({});
  const [topFilters, setTopFilters] = useState<Record<string, string>>({});
  const [advancedTableFiltersExpanded, setAdvancedTableFiltersExpanded] = useState(false);
  const [selectedViewModeId, setSelectedViewModeId] = useState<string>(
    () => listView?.activeViewModeId ?? listView?.viewModes?.[0]?.id ?? 'all'
  );
  const [fieldMetadata, setFieldMetadata] = useState<Awaited<ReturnType<FieldsService['getVisibleFields']>> | undefined>(undefined);
  const [dynamicContext, setDynamicContext] = useState<IDynamicContext | undefined>(undefined);
  const [listDisplayMode, setListDisplayMode] = useState<'table' | 'cards'>(() =>
    listView?.listCardViewEnabled === true && listView?.listDefaultDisplayMode === 'cards' ? 'cards' : 'table'
  );
  const listCardViewEnabled = listView?.listCardViewEnabled === true;

  const fullViewModes = listView?.viewModes ?? EMPTY_VIEW_MODES;
  const membership = useViewModeMembership(fullViewModes, pageWebServerRelativeUrl);
  const visibleViewModes = useMemo(() => {
    if (!membership) return fullViewModes;
    return filterViewModesForCurrentUser(
      fullViewModes,
      membership.userId,
      membership.groupByWeb,
      membership.pageNorm
    );
  }, [fullViewModes, membership]);

  // Reseta o modo somente quando a lista muda (novo listTitle), não por mudança de referência da config.
  const prevListTitleRef = useRef(listTitle);
  if (prevListTitleRef.current !== listTitle) {
    prevListTitleRef.current = listTitle;
    setSelectedViewModeId(listView?.activeViewModeId ?? fullViewModes[0]?.id ?? 'all');
  }

  // Quando membership chega pela primeira vez (fetch assíncrono), valida se o modo ainda é permitido.
  const membershipInitializedRef = useRef(false);
  useEffect(() => {
    if (!membership) return;
    if (membershipInitializedRef.current) return;
    membershipInitializedRef.current = true;
    const modes = fullViewModes;
    const visible = filterViewModesForCurrentUser(
      modes,
      membership.userId,
      membership.groupByWeb,
      membership.pageNorm
    );
    setSelectedViewModeId((prev) =>
      visible.some((m) => m.id === prev)
        ? prev
        : pickFallbackViewModeId(listView?.activeViewModeId ?? modes[0]?.id, visible, modes)
    );
  // Dependência intencional: só deve rodar na primeira chegada do membership.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [membership]);

  useEffect(() => {
    setColumnFilters({});
  }, [selectedViewModeId]);

  useEffect(() => {
    if (clearFiltersSignal === undefined) return;
    setColumnFilters({});
    setTopFilters({});
  }, [clearFiltersSignal]);

  useEffect(() => {
    if (listView.sort?.field != null && String(listView.sort.field).trim()) {
      setSortConfig({
        field: String(listView.sort.field).trim(),
        direction: listView.sort.ascending === false ? 'desc' : 'asc',
      });
    } else {
      setSortConfig(null);
    }
  }, [listTitle, listView.sort?.field, listView.sort?.ascending]);

  const onActiveViewModeChangeRef = useRef(onActiveViewModeChange);
  onActiveViewModeChangeRef.current = onActiveViewModeChange;

  useEffect(() => {
    onActiveViewModeChangeRef.current?.(selectedViewModeId);
  }, [selectedViewModeId]);

  useEffect(() => {
    if (!listCardViewEnabled) {
      setListDisplayMode('table');
      return;
    }
    setListDisplayMode(listView?.listDefaultDisplayMode === 'cards' ? 'cards' : 'table');
  }, [listCardViewEnabled, listView?.listDefaultDisplayMode]);

  useEffect(() => {
    const usersService = new UsersService();
    usersService
      .getCurrentUser()
      .then((user) => {
        setDynamicContext(
          buildDynamicContext({
            currentUser: { id: user.Id, title: user.Title, name: user.Title, email: user.Email, loginName: user.LoginName },
            query: typeof window !== 'undefined' && window.location ? parseQueryString(window.location.search) : undefined,
            now: new Date(),
            site:
              pageWebServerRelativeUrl?.trim().length
                ? { url: pageWebServerRelativeUrl.trim() }
                : undefined,
          })
        );
      })
      .catch(() =>
        setDynamicContext(
          buildDynamicContext({
            now: new Date(),
            site:
              pageWebServerRelativeUrl?.trim().length
                ? { url: pageWebServerRelativeUrl.trim() }
                : undefined,
          })
        )
      );
  }, [pageWebServerRelativeUrl]);

  function buildColumnFilterString(filters: Record<string, string>): string | undefined {
    const parts: string[] = [];
    for (const field in filters) {
      if (Object.prototype.hasOwnProperty.call(filters, field)) {
        const val = (filters[field] || '').trim();
        if (val) parts.push(`substringof('${String(val).replace(/'/g, "''")}', ${field})`);
      }
    }
    return parts.length === 0 ? undefined : parts.join(' and ');
  }

  const engine = useMemo(() => new TableEngine(), []);
  const itemsService = useMemo(() => new ItemsService(), []);
  const fieldsService = useMemo(() => new FieldsService(), []);

  useEffect(() => {
    if (!listTitle.trim()) return;
    setFieldMetadata(undefined);
    fieldsService
      .getVisibleFields(listTitle, listWeb)
      .then((f) => {
        const extra = SYSTEM_METADATA_FIELDS.filter(
          (sf) => !f.some((x) => x.InternalName === sf.InternalName)
        );
        setFieldMetadata([...f, ...extra]);
      })
      .catch(() => setFieldMetadata([]));
  }, [listTitle, listWeb]);

  useEffect(() => {
    if (!fieldMetadata) return;
    const normalized = engine.normalizeTableConfig(initialTableConfig, fieldMetadata);
    setTableConfig(normalized);
    setSortConfig((prev) => {
      if (prev?.field) {
        const field = prev.field;
        for (let i = 0; i < normalized.columns.length; i++) {
          const c = normalized.columns[i];
          const path = columnODataPath(c);
          const prefix = c.internalName + '/';
          if (field === path || field === c.internalName || field.indexOf(prefix) === 0) {
            if (!c.sortable) return normalized.defaultSort ?? null;
            break;
          }
        }
      }
      if (prev) return prev;
      return normalized.defaultSort ?? null;
    });
  }, [fieldMetadata, initialTableConfig]);

  const effectiveSort = sortConfig ?? tableConfig?.defaultSort ?? null;
  const pageSize = pagination?.enabled ? pagination.pageSize : 100;

  const topFiltersOData = useMemo(
    () => buildTableTopFiltersOData(topFilters, fieldMetadata ?? []),
    [topFilters, fieldMetadata]
  );

  const visibleColumns = useMemo(() => {
    if (!tableConfig) return [];
    return engine.getVisibleColumns(tableConfig);
  }, [engine, tableConfig]);

  const displayColumns = useResponsiveListTableColumns(visibleColumns);

  const pagingResetKey = useMemo(() => {
    if (!tableConfig) return `pending|${listTitle}|${listWeb ?? ''}`;
    const columns = engine.getVisibleColumns(tableConfig);
    if (columns.length === 0) return `empty|${listTitle}|${listWeb ?? ''}`;
    const columnFilterStr = buildColumnFilterString(columnFilters);
    const listViewWithMode = { ...listView, activeViewModeId: selectedViewModeId };
    const viewModeFilters = getActiveViewModeFilters(listViewWithMode);
    const viewModeFilterStr = buildListFilter(viewModeFilters, { dynamicContext, fieldsMetadata: fieldMetadata });
    const dashboardFilterStr =
      dashboardListFilters && dashboardListFilters.length > 0
        ? buildListFilter(dashboardListFilters, { dynamicContext, fieldsMetadata: fieldMetadata })
        : undefined;
    const sortPart = `${effectiveSort?.field ?? ''}|${effectiveSort?.direction ?? ''}`;
    const colsKey = columns.map((c) => `${c.internalName}/${c.expandConfig?.displayField ?? ''}`).join(',');
    return [
      listTitle,
      listWeb ?? '',
      String(pageSize),
      sortPart,
      colsKey,
      viewModeFilterStr ?? '',
      dashboardFilterStr ?? '',
      columnFilterStr ?? '',
      topFiltersOData ?? '',
      String(fieldMetadata?.length ?? 0),
    ].join('||');
  }, [
    listTitle,
    listWeb,
    pageSize,
    effectiveSort?.field,
    effectiveSort?.direction,
    tableConfig,
    listView,
    selectedViewModeId,
    dynamicContext,
    fieldMetadata,
    dashboardListFilters,
    columnFilters,
    topFiltersOData,
    engine,
  ]);

  useEffect(() => {
    if (!listTitle.trim() || !tableConfig) return;
    const columns = engine.getVisibleColumns(tableConfig);
    if (columns.length === 0) {
      setLoading(false);
      return;
    }

    if (paging.resetKey !== pagingResetKey) {
      setPaging({ pageIndex: 0, forwardPivots: [], resetKey: pagingResetKey });
      return;
    }

    setLoading(true);
    setError(undefined);
    const columnFilterStr = buildColumnFilterString(columnFilters);
    const listViewWithMode = { ...listView, activeViewModeId: selectedViewModeId };
    const viewModeFilters = getActiveViewModeFilters(listViewWithMode);
    const viewModeFilterStr = buildListFilter(viewModeFilters, { dynamicContext, fieldsMetadata: fieldMetadata });
    const dashboardFilterStr =
      dashboardListFilters && dashboardListFilters.length > 0
        ? buildListFilter(dashboardListFilters, { dynamicContext, fieldsMetadata: fieldMetadata })
        : undefined;
    const filterParts = [viewModeFilterStr, dashboardFilterStr, columnFilterStr, topFiltersOData].filter(Boolean);
    const combinedFilter = filterParts.length > 0 ? filterParts.join(' and ') : undefined;
    const request = engine.buildDataRequest({
      sortConfig: effectiveSort,
      top: pageSize,
      filter: combinedFilter,
    });

    const options = {
      select: request.select,
      expand: request.expand,
      orderBy: request.orderBy,
      filter: request.filter,
      fieldMetadata,
      ...(listWeb ? { webServerRelativeUrl: listWeb } : {}),
    };

    const afterLastItemId =
      paging.pageIndex === 0 ? undefined : paging.forwardPivots[paging.pageIndex - 1];

    let cancelled = false;
    itemsService
      .getPagedItems<Record<string, unknown>>(listTitle, options, pageSize, afterLastItemId)
      .then(
        (result) => {
          if (cancelled) return;
          setItems(result.items);
          setHasNext(result.hasNext);
          setLoading(false);
        },
        (err: Error) => {
          if (cancelled) return;
          setError(err.message);
          setItems([]);
          setLoading(false);
        }
      );
    return () => {
      cancelled = true;
    };
  }, [
    itemsService,
    listTitle,
    listWeb,
    tableConfig,
    effectiveSort,
    pageSize,
    columnFilters,
    topFiltersOData,
    selectedViewModeId,
    listView,
    fieldMetadata,
    dynamicContext,
    dashboardListFilters,
    pagingResetKey,
    paging,
    engine,
  ]);

  const handleSort = (field: string, direction: 'asc' | 'desc'): void => {
    setSortConfig({ field, direction });
  };

  const handleColumnFilter = (field: string, value: string): void => {
    setColumnFilters((prev) => {
      const next = { ...prev };
      if ((value || '').trim()) next[field] = value.trim();
      else delete next[field];
      return next;
    });
  };

  const layout = pagination?.layout ?? 'buttons';
  const currentPage = paging.pageIndex + 1;
  const from = paging.pageIndex * pageSize + 1;
  const to = paging.pageIndex * pageSize + items.length;
  const showPagination = pagination?.enabled && (hasNext || paging.pageIndex > 0);
  const onPrev = (): void => {
    setPaging((prev) =>
      prev.resetKey !== pagingResetKey || prev.pageIndex <= 0
        ? prev
        : { ...prev, pageIndex: prev.pageIndex - 1 }
    );
  };
  const onNext = (): void => {
    const last = readListItemId(items[items.length - 1]);
    if (last === undefined) return;
    setPaging((prev) => {
      if (prev.resetKey !== pagingResetKey) return prev;
      const pivots = prev.forwardPivots.slice(0, prev.pageIndex);
      pivots[prev.pageIndex] = last;
      return { ...prev, pageIndex: prev.pageIndex + 1, forwardPivots: pivots };
    });
  };
  const goToPage = (page: number): void => {
    setPaging((prev) =>
      prev.resetKey !== pagingResetKey ? prev : { ...prev, pageIndex: Math.max(0, page - 1) }
    );
  };

  const pagedNumbers: (number | 'ellipsis')[] =
    layout === 'paged'
      ? currentPage <= 3
        ? (() => { const a: number[] = []; for (let i = 1; i <= currentPage; i++) a.push(i); return a; })()
        : [1, 'ellipsis', currentPage - 2, currentPage - 1, currentPage]
      : [];

  const paginationBar =
    showPagination && (
      <Stack
        className={`${DINAMIC_SX_TABLE_CLASS.pagination}${layout === 'compact' ? ' dinamicSxTablePagination--compact' : ''}`}
        horizontal
        tokens={{ childrenGap: 8 }}
        horizontalAlign="end"
        styles={{ root: { flexWrap: 'wrap' } }}
      >
        {layout === 'compact' && (
          <span style={{ alignSelf: 'center', marginRight: 8, fontSize: 12 }}>
            {from}–{to}
          </span>
        )}
        {layout === 'numbered' && (
          <span style={{ alignSelf: 'center', marginRight: 8, fontSize: 12 }}>
            Página {currentPage}
          </span>
        )}
        {layout === 'paged' && (
          <>
            {paging.pageIndex > 0 && (
              <button type="button" className={DINAMIC_SX_TABLE_CLASS.paginationBtn} onClick={onPrev}>
                Anterior
              </button>
            )}
            {pagedNumbers.map((n, i) =>
              n === 'ellipsis' ? (
                <span key={`e-${i}`} style={{ alignSelf: 'center', padding: '0 4px' }}>
                  …
                </span>
              ) : (
                <button
                  key={n}
                  type="button"
                  className={DINAMIC_SX_TABLE_CLASS.paginationBtn}
                  data-active={n === currentPage ? 'true' : undefined}
                  onClick={() => goToPage(n)}
                >
                  {n}
                </button>
              )
            )}
            {hasNext && (
              <button type="button" className={DINAMIC_SX_TABLE_CLASS.paginationBtn} onClick={onNext}>
                Próxima
              </button>
            )}
          </>
        )}
        {layout !== 'paged' && (
          <>
            {paging.pageIndex > 0 && (
              <button type="button" className={DINAMIC_SX_TABLE_CLASS.paginationBtn} onClick={onPrev}>
                {layout === 'compact' ? '‹' : 'Anterior'}
              </button>
            )}
            {hasNext && (
              <button type="button" className={DINAMIC_SX_TABLE_CLASS.paginationBtn} onClick={onNext}>
                {layout === 'compact' ? '›' : 'Próxima'}
              </button>
            )}
          </>
        )}
      </Stack>
    );

  const viewModeOptions: IDropdownOption[] = visibleViewModes.map((m) => ({ key: m.id, text: m.label }));

  const tableFilterFieldsMetaSplit = useMemo(() => {
    type Row = {
      config: NonNullable<IListViewConfig['tableFilterFields']>[number];
      meta: import('../../../../services/shared/types').IFieldMetadata | null;
    };
    const empty: { fixed: Row[]; advanced: Row[] } = { fixed: [], advanced: [] };
    if (!listView?.tableFilterFields?.length || !fieldMetadata?.length) return empty;
    const metaByName = new Map((fieldMetadata as import('../../../../services/shared/types').IFieldMetadata[]).map((m) => [m.InternalName, m]));
    const fixed: Row[] = [];
    const advanced: Row[] = [];
    for (let i = 0; i < listView.tableFilterFields.length; i++) {
      const f = listView.tableFilterFields[i];
      const baseName = f.field.indexOf('/') !== -1 ? f.field.split('/')[0] : f.field;
      const meta = metaByName.get(baseName) ?? null;
      const row: Row = { config: f, meta };
      if (f.placement === 'advanced') advanced.push(row);
      else fixed.push(row);
    }
    return { fixed, advanced };
  }, [listView?.tableFilterFields, fieldMetadata]);

  const hasTopFilters =
    tableFilterFieldsMetaSplit.fixed.length > 0 || tableFilterFieldsMetaSplit.advanced.length > 0;

  const chromeBySlot = useMemo(() => {
    const m = new Map<TListViewChromeButtonSlot, IListViewChromeButtonConfig[]>();
    for (const b of listView?.chromeButtons ?? []) {
      const arr = m.get(b.slot) ?? [];
      arr.push(b);
      m.set(b.slot, arr);
    }
    return m;
  }, [listView?.chromeButtons]);

  const hasChromeToolbarSlot = useMemo(() => {
    const toolbarSlots: TListViewChromeButtonSlot[] = [
      'toolbarAfterViewMode',
      'toolbarAfterTableCardsToggle',
      'toolbarAfterPdfExport',
      'toolbarBeforeClearFilters',
    ];
    for (let i = 0; i < toolbarSlots.length; i++) {
      if ((chromeBySlot.get(toolbarSlots[i])?.length ?? 0) > 0) return true;
    }
    return false;
  }, [chromeBySlot]);

  const chromeFiltersAfterToggle = useMemo(
    () => sortChromeForSlot(chromeBySlot.get('filtersAfterAdvancedToggle') ?? []),
    [chromeBySlot]
  );
  const chromeFiltersBelow = useMemo(
    () => sortChromeForSlot(chromeBySlot.get('filtersBelowControls') ?? []),
    [chromeBySlot]
  );

  const showFilterBar =
    hasTopFilters || chromeFiltersAfterToggle.length > 0 || chromeFiltersBelow.length > 0;

  const advancedTableFiltersTitle =
    listView?.tableAdvancedFiltersTitle?.trim() || 'Filtros avançados';

  const activeTopFiltersCount = Object.values(topFilters).filter((v) => v.trim()).length;
  const hasActiveColumnFilters = Object.values(columnFilters).some((v) => v.trim().length > 0);
  const hasAnyActiveFilter =
    hasActiveColumnFilters ||
    activeTopFiltersCount > 0 ||
    (dashboardListFilters?.length ?? 0) > 0;

  const handleClearAllFilters = (): void => {
    setColumnFilters({});
    setTopFilters({});
    onClearFilters?.();
  };

  const renderTopFilterControl = (fieldCfg: { config: { field: string; label?: string }; meta: import('../../../../services/shared/types').IFieldMetadata | null }): React.ReactNode => {
    const { config: fc, meta } = fieldCfg;
    const label = (fc.label && fc.label.trim()) ? fc.label.trim() : meta?.Title || fc.field;
    const val = topFilters[fc.field] ?? '';
    const onChange = (v: string): void =>
      setTopFilters((prev) => {
        if (!v.trim()) {
          const next = { ...prev };
          delete next[fc.field];
          return next;
        }
        return { ...prev, [fc.field]: v };
      });
    const mtype = meta?.MappedType ?? 'text';

    if (mtype === 'choice' || mtype === 'multichoice') {
      const choiceOptions: IDropdownOption[] = [
        { key: '', text: `Todos` },
        ...(meta?.Choices ?? []).map((c) => ({ key: c, text: c })),
      ];
      return (
        <div key={fc.field} className={DINAMIC_SX_FILTER_CLASS.control}>
          <Dropdown
            label={label}
            selectedKey={val}
            options={choiceOptions}
            onChange={(_, opt) => onChange(opt?.key === '' ? '' : String(opt?.key ?? ''))}
            styles={{ root: { display: 'block', margin: 0 } }}
          />
        </div>
      );
    }
    if (mtype === 'boolean') {
      const boolOptions: IDropdownOption[] = [
        { key: '', text: 'Todos' },
        { key: 'true', text: 'Sim' },
        { key: 'false', text: 'Não' },
      ];
      return (
        <div key={fc.field} className={DINAMIC_SX_FILTER_CLASS.control}>
          <Dropdown
            label={label}
            selectedKey={val}
            options={boolOptions}
            onChange={(_, opt) => onChange(opt?.key === '' ? '' : String(opt?.key ?? ''))}
            styles={{ root: { display: 'block', margin: 0 } }}
          />
        </div>
      );
    }
    if (mtype === 'datetime') {
      return (
        <div key={fc.field} className={DINAMIC_SX_FILTER_CLASS.control}>
          <label className={DINAMIC_SX_FILTER_CLASS.label} htmlFor={`filter-${fc.field}`}>
            {label}
          </label>
          <input
            id={`filter-${fc.field}`}
            type="date"
            className={DINAMIC_SX_FILTER_CLASS.input}
            value={val}
            onChange={(e) => onChange(e.target.value)}
            aria-label={label}
          />
        </div>
      );
    }
    return (
      <div key={fc.field} className={DINAMIC_SX_FILTER_CLASS.control}>
        <TextField
          label={label}
          value={val}
          onChange={(_, v) => onChange(v ?? '')}
          placeholder="Filtrar…"
          styles={{ root: { margin: 0 } }}
        />
      </div>
    );
  };

  const mergedTableCss = resolveTableLayoutCss(listView?.customTableCssSlots, listView?.customTableCss);
  const rowRulesCss = mergeRowStyleRulesCss(listView?.tableRowStyleRules);
  const instanceScopeClass = `dinamicSxScope_${instanceScopeId.replace(/[^a-zA-Z0-9_-]/g, '_')}`;
  const mergedLayoutCssRaw = [mergedTableCss, rowRulesCss].filter((s) => s.length > 0).join('\n\n').trim();
  const mergedLayoutCss = [
    scopeTableCssByInstance(mergedLayoutCssRaw, instanceScopeClass),
    TABLE_COLUMN_FILTER_PORTAL_CSS,
  ]
    .filter(Boolean)
    .join('\n\n');
  const mergedCardCss = scopeCardCssByInstance(listView?.customCardCss ?? '', instanceScopeClass);
  const mergedFilterCss = scopeFilterCssByInstance(
    resolveFilterBarCss(listView?.customFilterCss),
    instanceScopeClass
  );
  const mergedViewModeCss = scopeViewModeCssByInstance(
    resolveViewModeCss(listView?.customViewModeCss),
    instanceScopeClass
  );
  const mergedToolbarCss = scopeToolbarCssByInstance(resolveListToolbarCss(), instanceScopeClass);
  const tableCustomStyle =
    mergedLayoutCss.length > 0 ||
    mergedCardCss.length > 0 ||
    mergedFilterCss.length > 0 ||
    mergedViewModeCss.length > 0 ||
    mergedToolbarCss.length > 0
      ? (
          <style type="text/css">
            {[mergedLayoutCss, mergedCardCss, mergedFilterCss, mergedViewModeCss, mergedToolbarCss]
              .filter(Boolean)
              .join('\n\n')}
          </style>
        )
      : null;

  const actionContext = dynamicContext ?? { now: new Date() };
  const chromeRowContext: Record<string, unknown> =
    items.length > 0 ? (items[0] as Record<string, unknown>) : {};
  const listRowActions = listView?.listRowActions;
  const userGroupIds: Set<number> | undefined = membership?.groupByWeb?.get(membership.pageNorm) ?? (membership ? new Set<number>() : undefined);

  if (!tableConfig) {
    return (
      <>
        {tableCustomStyle}
        <DataTable
          config={{ enabled: true, columns: [], sortable: false, emptyMessage: '' }}
          displayColumns={displayColumns}
          items={[]}
          loading={true}
          sortConfig={null}
          onSort={handleSort}
          engine={engine}
          rowStyleRules={listView?.tableRowStyleRules}
          rowActions={listRowActions}
          dynamicContext={actionContext}
          userGroupIds={userGroupIds}
        />
      </>
    );
  }

  const showPdfButton = listView?.pdfExportEnabled === true;

  const handleExportPdf = async (): Promise<void> => {
    const template = config.pdfTemplate;
    if (!template?.body?.elements?.length) return;
    const data = items as Record<string, unknown>[];
    if (data.length === 0) return;
    const name = `${dataSource.title || 'lista'}_${new Date().toISOString().slice(0, 10)}.pdf`;
    await generateAndDownloadPdf(template, data, name);
  };

  const renderListChromeButton = (
    it: IListPageButtonItemConfig,
    alignSelf?: 'flex-end'
  ): React.ReactNode => {
    const btnStyle = parseListChromeCss(it.css);
    const iconProps = it.iconName ? { iconName: it.iconName } : undefined;
    const className =
      it.variant === 'primary' ? DINAMIC_SX_TOOLBAR_CLASS.primaryBtn : DINAMIC_SX_TOOLBAR_CLASS.defaultBtn;
    const btn =
      it.variant === 'primary' ? (
        <PrimaryButton
          className={className}
          text={it.label}
          iconProps={iconProps}
          onClick={() => navigateListChromeButton(it, actionContext, chromeRowContext)}
        />
      ) : (
        <DefaultButton
          className={className}
          text={it.label}
          iconProps={iconProps}
          onClick={() => navigateListChromeButton(it, actionContext, chromeRowContext)}
        />
      );
    const wrapStyle = alignSelf ? { ...btnStyle, alignSelf } : btnStyle;
    return wrapStyle ? (
      <span style={wrapStyle}>{btn}</span>
    ) : (
      btn
    );
  };

  const renderToolbarChrome = (slot: TListViewChromeButtonSlot): React.ReactNode => {
    const sorted = sortChromeForSlot(chromeBySlot.get(slot) ?? []);
    if (!sorted.length) return null;
    return (
      <>
        {sorted.map((it) => (
          <React.Fragment key={it.id}>{renderListChromeButton(it)}</React.Fragment>
        ))}
      </>
    );
  };

  const renderInlineFilterChrome = (): React.ReactNode =>
    chromeFiltersAfterToggle.map((it) => (
      <React.Fragment key={it.id}>{renderListChromeButton(it, 'flex-end')}</React.Fragment>
    ));

  const renderBelowFilterChrome = (): React.ReactNode =>
    chromeFiltersBelow.map((it) => (
      <React.Fragment key={it.id}>{renderListChromeButton(it)}</React.Fragment>
    ));

  const showChromeRow =
    viewModeOptions.length > 0 ||
    showPdfButton ||
    listCardViewEnabled ||
    hasAnyActiveFilter ||
    hasChromeToolbarSlot ||
    showFilterBar;

  const hasFixedFilterFields = tableFilterFieldsMetaSplit.fixed.length > 0;
  const showAdvancedFilterPanel =
    hasTopFilters &&
    tableFilterFieldsMetaSplit.advanced.length > 0 &&
    advancedTableFiltersExpanded;
  const showFilterFieldsPanel =
    hasFixedFilterFields || showAdvancedFilterPanel || chromeFiltersBelow.length > 0;
  const advancedOnlyPanel =
    showAdvancedFilterPanel && !hasFixedFilterFields && chromeFiltersBelow.length === 0;

  const renderListChromeRow = (): React.ReactNode => (
    <div className={DINAMIC_SX_TOOLBAR_CLASS.chromeRow}>
      <div className={DINAMIC_SX_TOOLBAR_CLASS.chromeRowStart}>
        {viewModeOptions.length > 0 ? (
          <ViewModePickerBar
            picker={listView?.viewModePicker}
            modes={visibleViewModes}
            selectedId={selectedViewModeId}
            onSelect={setSelectedViewModeId}
            options={viewModeOptions}
          />
        ) : null}
      </div>
      <div className={DINAMIC_SX_TOOLBAR_CLASS.chromeRowEnd}>
        {renderToolbarChrome('toolbarAfterViewMode')}
        {hasTopFilters && tableFilterFieldsMetaSplit.advanced.length > 0 ? (
          <ActionButton
            className={`${DINAMIC_SX_FILTER_CLASS.advancedBtn}${advancedTableFiltersExpanded ? ' dinamicSxFilterAdvancedBtn--open' : ''}`}
            iconProps={{
              iconName: advancedTableFiltersExpanded ? 'ChevronUp' : 'ChevronDown',
            }}
            onClick={() => setAdvancedTableFiltersExpanded((x) => !x)}
            aria-expanded={advancedTableFiltersExpanded}
          >
            {advancedTableFiltersTitle}
          </ActionButton>
        ) : null}
        {renderInlineFilterChrome()}
        {activeTopFiltersCount > 0 ? (
          <ActionButton
            className={DINAMIC_SX_FILTER_CLASS.headerClear}
            iconProps={{ iconName: 'ClearFilter' }}
            text="Limpar"
            onClick={() => setTopFilters({})}
          />
        ) : null}
        {listCardViewEnabled ? (
          <>
            <TableCardsLayoutToggle value={listDisplayMode} onChange={setListDisplayMode} />
            {renderToolbarChrome('toolbarAfterTableCardsToggle')}
          </>
        ) : null}
        {showPdfButton ? (
          <ActionButton
            className={DINAMIC_SX_TOOLBAR_CLASS.ghostBtn}
            iconProps={{ iconName: 'PDF' }}
            text="Exportar PDF"
            onClick={handleExportPdf}
          />
        ) : null}
        {renderToolbarChrome('toolbarAfterPdfExport')}
        {renderToolbarChrome('toolbarBeforeClearFilters')}
        {hasAnyActiveFilter ? (
          <ActionButton
            className={DINAMIC_SX_TOOLBAR_CLASS.ghostBtn}
            iconProps={{ iconName: 'ClearFilter' }}
            onClick={handleClearAllFilters}
          >
            Remover Filtros
          </ActionButton>
        ) : null}
      </div>
    </div>
  );

  return (
    <Stack
      className={`${instanceScopeClass} ${DINAMIC_SX_TABLE_CLASS.viewRoot}`}
      tokens={{ childrenGap: 12 }}
      styles={{ root: { marginTop: 8 } }}
    >
      {tableCustomStyle}
      <Stack className="dinamicSxFilterTableBlock" tokens={{ childrenGap: 15 }}>
        {showChromeRow ? renderListChromeRow() : null}
        {showFilterFieldsPanel ? (
          <Stack
            className={`${DINAMIC_SX_FILTER_CLASS.bar}${
              advancedOnlyPanel ? ` ${DINAMIC_SX_FILTER_CLASS.barAdvancedOnly}` : ''
            }${
              hasFixedFilterFields && !showAdvancedFilterPanel && chromeFiltersBelow.length === 0
                ? ` ${DINAMIC_SX_FILTER_CLASS.barFieldsOnly}`
                : ''
            }`}
            tokens={{ childrenGap: 0 }}
          >
            {hasFixedFilterFields ? (
              <div className={DINAMIC_SX_FILTER_CLASS.fieldsRow}>
                {tableFilterFieldsMetaSplit.fixed.map((f) => renderTopFilterControl(f))}
              </div>
            ) : null}
            {showAdvancedFilterPanel ? (
              <div className={DINAMIC_SX_FILTER_CLASS.advancedPanel}>
                <div className={DINAMIC_SX_FILTER_CLASS.fieldsRow}>
                  {tableFilterFieldsMetaSplit.advanced.map((f) => renderTopFilterControl(f))}
                </div>
              </div>
            ) : null}
            {chromeFiltersBelow.length > 0 ? (
              <div className={DINAMIC_SX_FILTER_CLASS.advancedPanel}>
                <div className={DINAMIC_SX_FILTER_CLASS.fieldsRow}>{renderBelowFilterChrome()}</div>
              </div>
            ) : null}
          </Stack>
        ) : null}
        {listDisplayMode === 'cards' && listCardViewEnabled ? (
          <ListItemsCardGrid
            columns={engine.getVisibleColumns(tableConfig)}
            displayColumns={displayColumns}
            items={items}
            loading={loading}
            error={error}
            emptyMessage={tableConfig.emptyMessage ?? 'Nenhum item encontrado.'}
            engine={engine}
            sortConfig={effectiveSort}
            onSort={handleSort}
            tableSortable={tableConfig.sortable}
            columnFilters={columnFilters}
            onColumnFilter={handleColumnFilter}
            dense={tableConfig.dense}
            rowActions={listRowActions}
            dynamicContext={actionContext}
            userGroupIds={userGroupIds}
          />
        ) : (
          <DataTable
            config={tableConfig}
            displayColumns={displayColumns}
            items={items}
            loading={loading}
            error={error}
            sortConfig={effectiveSort}
            onSort={handleSort}
            columnFilters={columnFilters}
            onColumnFilter={handleColumnFilter}
            engine={engine}
            rowStyleRules={listView?.tableRowStyleRules}
            rowActions={listRowActions}
            dynamicContext={actionContext}
            userGroupIds={userGroupIds}
          />
        )}
      </Stack>
      {paginationBar}
    </Stack>
  );
};
