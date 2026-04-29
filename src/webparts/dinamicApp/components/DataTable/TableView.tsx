import * as React from 'react';
import { useState, useEffect, useMemo, useRef } from 'react';
import {
  Stack,
  Text,
  Dropdown,
  IDropdownOption,
  ActionButton,
  ChoiceGroup,
  IChoiceGroupOption,
  DefaultButton,
  TextField,
} from '@fluentui/react';
import { IDynamicViewConfig, IListViewFilterConfig, IListViewModeConfig } from '../../core/config/types';
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
import { DINAMIC_SX_TABLE_CLASS, mergeCustomTableCss, mergeRowStyleRulesCss, scopeCardCssByInstance } from './tableLayoutClasses';
import type { IDynamicContext } from '../../core/dynamicTokens/types';

const EMPTY_VIEW_MODES: IListViewModeConfig[] = [];

function listViewToTableConfig(listView: IDynamicViewConfig['listView']): Partial<ITableConfig> {
  const columns = (listView.columns ?? []).map((c) => ({
    id: c.field,
    internalName: c.field,
    label: c.label ?? c.field,
    visible: true,
    sortable: true,
    expandConfig: c.expandField ? { displayField: c.expandField } : undefined,
  }));
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
}

function scopeTableCssByInstance(css: string, scopeClass: string): string {
  if (!css.trim()) return '';
  return css.replace(/\.dinamicSxTable/g, `.${scopeClass} .dinamicSxTable`);
}

function scopeFilterCssByInstance(css: string, scopeClass: string): string {
  if (!css.trim()) return '';
  return css.replace(/\.dinamicSxFilter/g, `.${scopeClass} .dinamicSxFilter`);
}

function scopeViewModeCssByInstance(css: string, scopeClass: string): string {
  if (!css.trim()) return '';
  return css.replace(/\.dinamicSxViewMode/g, `.${scopeClass} .dinamicSxViewMode`);
}

export const TableView: React.FC<ITableViewProps> = ({
  config,
  dashboardListFilters,
  instanceScopeId,
  pageWebServerRelativeUrl,
  onActiveViewModeChange,
  clearFiltersSignal,
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
          })
        );
      })
      .catch(() => setDynamicContext(buildDynamicContext({ now: new Date() })));
  }, []);

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
          const prefix = c.internalName + '/';
          if (field === c.internalName || field.indexOf(prefix) === 0) {
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
    const colsKey = columns.map((c) => c.internalName).join(',');
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

  const btnPad = layout === 'compact' ? '4px 8px' : '6px 12px';

  const paginationBar =
    showPagination && (
      <Stack
        className={DINAMIC_SX_TABLE_CLASS.pagination}
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
              <button type="button" onClick={onPrev} style={{ padding: btnPad, cursor: 'pointer' }}>
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
                  onClick={() => goToPage(n)}
                  style={{
                    padding: btnPad,
                    cursor: 'pointer',
                    fontWeight: n === currentPage ? 'bold' : undefined,
                  }}
                >
                  {n}
                </button>
              )
            )}
            {hasNext && (
              <button type="button" onClick={onNext} style={{ padding: btnPad, cursor: 'pointer' }}>
                Próxima
              </button>
            )}
          </>
        )}
        {layout !== 'paged' && (
          <>
            {paging.pageIndex > 0 && (
              <button type="button" onClick={onPrev} style={{ padding: btnPad, cursor: 'pointer' }}>
                {layout === 'compact' ? '‹' : 'Anterior'}
              </button>
            )}
            {hasNext && (
              <button type="button" onClick={onNext} style={{ padding: btnPad, cursor: 'pointer' }}>
                {layout === 'compact' ? '›' : 'Próxima'}
              </button>
            )}
          </>
        )}
      </Stack>
    );

  const viewModeOptions: IDropdownOption[] = visibleViewModes.map((m) => ({ key: m.id, text: m.label }));
  const viewModesAsTabs = listView?.viewModePicker === 'tabs';

  const tableFilterFieldsMeta = useMemo(() => {
    if (!listView?.tableFilterFields?.length || !fieldMetadata?.length) return [];
    const metaByName = new Map((fieldMetadata as import('../../../../services/shared/types').IFieldMetadata[]).map((m) => [m.InternalName, m]));
    return listView.tableFilterFields.map((f) => {
      const baseName = f.field.indexOf('/') !== -1 ? f.field.split('/')[0] : f.field;
      const meta = metaByName.get(baseName) ?? null;
      return { config: f, meta };
    });
  }, [listView?.tableFilterFields, fieldMetadata]);

  const hasTopFilters = tableFilterFieldsMeta.length > 0;

  const activeTopFiltersCount = Object.values(topFilters).filter((v) => v.trim()).length;

  const renderTopFilterControl = (fieldCfg: { config: { field: string; label?: string }; meta: import('../../../../services/shared/types').IFieldMetadata | null }): React.ReactNode => {
    const { config: fc, meta } = fieldCfg;
    const label = fc.label || meta?.Title || fc.field;
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

    const wrapperStyle: React.CSSProperties = {
      display: 'flex',
      flexDirection: 'column',
      justifyContent: 'flex-start',
    };

    if (mtype === 'choice' || mtype === 'multichoice') {
      const choiceOptions: IDropdownOption[] = [
        { key: '', text: `Todos` },
        ...(meta?.Choices ?? []).map((c) => ({ key: c, text: c })),
      ];
      return (
        <div key={fc.field} className="dinamicSxFilterControl" style={{ ...wrapperStyle, minWidth: 160, maxWidth: 240 }}>
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
        <div key={fc.field} className="dinamicSxFilterControl" style={{ ...wrapperStyle, minWidth: 120, maxWidth: 180 }}>
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
        <div key={fc.field} className="dinamicSxFilterControl" style={{ ...wrapperStyle, minWidth: 150, maxWidth: 220 }}>
          <label style={{ fontSize: 14, fontWeight: 600, color: '#323130', display: 'block', padding: '5px 0' }}>
            {label}
          </label>
          <input
            type="date"
            value={val}
            onChange={(e) => onChange(e.target.value)}
            style={{
              height: 32,
              border: '1px solid #605e5c',
              borderRadius: 2,
              padding: '0 8px',
              fontSize: 14,
              fontFamily: 'inherit',
              color: '#323130',
              background: '#fff',
              width: '100%',
              boxSizing: 'border-box',
            }}
          />
        </div>
      );
    }
    return (
      <div key={fc.field} className="dinamicSxFilterControl" style={{ ...wrapperStyle, minWidth: 140, maxWidth: 220 }}>
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

  const listPresentationOptions: IChoiceGroupOption[] = useMemo(
    () => [
      { key: 'table', text: 'Tabela', iconProps: { iconName: 'Table' } },
      { key: 'cards', text: 'Cards', iconProps: { iconName: 'Tiles' } },
    ],
    []
  );

  const mergedTableCss = mergeCustomTableCss(listView?.customTableCssSlots, listView?.customTableCss);
  const rowRulesCss = mergeRowStyleRulesCss(listView?.tableRowStyleRules);
  const instanceScopeClass = `dinamicSxScope_${instanceScopeId.replace(/[^a-zA-Z0-9_-]/g, '_')}`;
  const mergedLayoutCssRaw = [mergedTableCss, rowRulesCss].filter((s) => s.length > 0).join('\n\n').trim();
  const mergedLayoutCss = scopeTableCssByInstance(mergedLayoutCssRaw, instanceScopeClass);
  const mergedCardCss = scopeCardCssByInstance(listView?.customCardCss ?? '', instanceScopeClass);
  const mergedFilterCss = scopeFilterCssByInstance(listView?.customFilterCss ?? '', instanceScopeClass);
  const mergedViewModeCss = scopeViewModeCssByInstance(listView?.customViewModeCss ?? '', instanceScopeClass);
  const tableCustomStyle =
    mergedLayoutCss.length > 0 || mergedCardCss.length > 0 || mergedFilterCss.length > 0 || mergedViewModeCss.length > 0
      ? <style type="text/css">{[mergedLayoutCss, mergedCardCss, mergedFilterCss, mergedViewModeCss].filter(Boolean).join('\n\n')}</style>
      : null;

  const actionContext = dynamicContext ?? { now: new Date() };
  const listRowActions = listView?.listRowActions;
  const userGroupIds: Set<number> | undefined = membership?.groupByWeb?.get(membership.pageNorm) ?? (membership ? new Set<number>() : undefined);

  if (!tableConfig) {
    return (
      <>
        {tableCustomStyle}
        <DataTable
          config={{ enabled: true, columns: [], sortable: false, emptyMessage: '' }}
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

  return (
    <Stack
      className={`${instanceScopeClass} ${DINAMIC_SX_TABLE_CLASS.viewRoot}`}
      tokens={{ childrenGap: 12 }}
      styles={{ root: { marginTop: 8 } }}
    >
      {tableCustomStyle}
      {(viewModeOptions.length > 0 || showPdfButton || listCardViewEnabled) && (
        <Stack
          className={DINAMIC_SX_TABLE_CLASS.toolbar}
          horizontal
          tokens={{ childrenGap: 12 }}
          verticalAlign="end"
          styles={{ root: { flexWrap: 'wrap' } }}
        >
          {viewModeOptions.length > 0 &&
            (viewModesAsTabs ? (
              <Stack className="dinamicSxViewModeBar" tokens={{ childrenGap: 4 }} styles={{ root: { flex: '1 1 auto', minWidth: 0 } }}>
                <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
                  Visualização
                </Text>
                <Stack
                  horizontal
                  wrap
                  tokens={{ childrenGap: 6 }}
                  verticalAlign="center"
                  role="tablist"
                  aria-label="Modos de visualização"
                  styles={{ root: { flexWrap: 'wrap' } }}
                >
                  {visibleViewModes.map((m) => (
                    <DefaultButton
                      key={m.id}
                      className="dinamicSxViewModeTab"
                      role="tab"
                      aria-selected={selectedViewModeId === m.id}
                      primary={selectedViewModeId === m.id}
                      text={m.label}
                      onClick={() => setSelectedViewModeId(m.id)}
                      styles={{ root: { minHeight: 32 } }}
                    />
                  ))}
                </Stack>
              </Stack>
            ) : (
              <div className="dinamicSxViewModeBar dinamicSxViewModeDropdown">
                <Dropdown
                  label="Visualização"
                  options={viewModeOptions}
                  selectedKey={selectedViewModeId}
                  onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                    if (opt) setSelectedViewModeId(String(opt.key));
                  }}
                  styles={{ root: { maxWidth: 220 } }}
                />
              </div>
            ))}
          {listCardViewEnabled && (
            <ChoiceGroup
              label="Apresentação"
              selectedKey={listDisplayMode}
              options={listPresentationOptions}
              onChange={(_, opt) => opt && setListDisplayMode(opt.key as 'table' | 'cards')}
              styles={{ flexContainer: { display: 'flex', flexWrap: 'wrap', columnGap: '12px', rowGap: '4px' } }}
            />
          )}
          {showPdfButton && (
            <ActionButton
              iconProps={{ iconName: 'PDF' }}
              text="Exportar PDF"
              styles={{ root: { height: 32, color: '#0078d4' } }}
              onClick={handleExportPdf}
            />
          )}
        </Stack>
      )}
      {hasTopFilters && (
        <Stack
          className="dinamicSxFilterBar"
          tokens={{ childrenGap: 8 }}
          styles={{
            root: {
              padding: '10px 14px',
              background: '#faf9f8',
              borderRadius: 8,
              border: '1px solid #edebe9',
            },
          }}
        >
          <Stack horizontal verticalAlign="center" horizontalAlign="space-between">
            <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
              Filtros{activeTopFiltersCount > 0 ? ` (${activeTopFiltersCount} ativo${activeTopFiltersCount > 1 ? 's' : ''})` : ''}
            </Text>
            {activeTopFiltersCount > 0 && (
              <ActionButton
                iconProps={{ iconName: 'ClearFilter' }}
                text="Limpar"
                styles={{ root: { height: 28, color: '#a4262c' } }}
                onClick={() => setTopFilters({})}
              />
            )}
          </Stack>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12, alignItems: 'flex-start' }}>
            {tableFilterFieldsMeta.map((f) => renderTopFilterControl(f))}
          </div>
        </Stack>
      )}
      {listDisplayMode === 'cards' && listCardViewEnabled ? (
        <ListItemsCardGrid
          columns={engine.getVisibleColumns(tableConfig)}
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
      {paginationBar}
    </Stack>
  );
};
