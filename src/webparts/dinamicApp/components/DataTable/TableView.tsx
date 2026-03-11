import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import { Stack, Dropdown, IDropdownOption } from '@fluentui/react';
import { IDynamicViewConfig } from '../../core/config/types';
import { TableEngine } from '../../core/table/services/TableEngine';
import type { ITableConfig, ISortConfig } from '../../core/table/types';
import { buildListFilter, getActiveViewModeFilters } from '../../core/listView';
import { ItemsService } from '../../../../services';
import { FieldsService } from '../../../../services';
import { DataTable } from './DataTable';
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
}

export const TableView: React.FC<ITableViewProps> = ({ config }) => {
  const { dataSource, pagination, listView, tableConfig: tableConfigRaw } = config;
  const listTitle = dataSource.title;

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
  const [skip, setSkip] = useState(0);
  const [hasNext, setHasNext] = useState(false);
  const [columnFilters, setColumnFilters] = useState<Record<string, string>>({});
  const [selectedViewModeId, setSelectedViewModeId] = useState<string>(
    () => listView?.activeViewModeId ?? listView?.viewModes?.[0]?.id ?? 'all'
  );
  const [fieldMetadata, setFieldMetadata] = useState<Awaited<ReturnType<FieldsService['getVisibleFields']>> | undefined>(undefined);

  useEffect(() => {
    setSelectedViewModeId(listView?.activeViewModeId ?? listView?.viewModes?.[0]?.id ?? 'all');
  }, [listView?.activeViewModeId, listView?.viewModes]);

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
    fieldsService.getVisibleFields(listTitle).then(setFieldMetadata).catch(() => setFieldMetadata([]));
  }, [listTitle]);

  useEffect(() => {
    if (!fieldMetadata) return;
    const normalized = engine.normalizeTableConfig(initialTableConfig, fieldMetadata);
    setTableConfig(normalized);
    if (normalized.defaultSort && !sortConfig) setSortConfig(normalized.defaultSort);
  }, [fieldMetadata, initialTableConfig]);

  const effectiveSort = sortConfig ?? tableConfig?.defaultSort ?? null;
  const pageSize = pagination?.enabled ? pagination.pageSize : 100;

  useEffect(() => {
    if (!listTitle.trim() || !tableConfig) return;
    const columns = engine.getVisibleColumns(tableConfig);
    if (columns.length === 0) {
      setLoading(false);
      return;
    }

    setLoading(true);
    setError(undefined);
    const columnFilterStr = buildColumnFilterString(columnFilters);
    const listViewWithMode = { ...listView, activeViewModeId: selectedViewModeId };
    const viewModeFilters = getActiveViewModeFilters(listViewWithMode);
    const viewModeFilterStr = buildListFilter(viewModeFilters);
    const filterParts = [viewModeFilterStr, columnFilterStr].filter(Boolean);
    const combinedFilter = filterParts.length > 0 ? filterParts.join(' and ') : undefined;
    const request = engine.buildDataRequest({
      sortConfig: effectiveSort,
      top: pageSize,
      skip,
      filter: combinedFilter,
    });

    const options = {
      select: request.select,
      expand: request.expand,
      orderBy: request.orderBy,
      filter: request.filter,
      fieldMetadata,
    };

    itemsService
      .getPagedItems<Record<string, unknown>>(listTitle, options, pageSize, skip)
      .then(
        (result) => {
          setItems(result.items);
          setHasNext(result.hasNext);
          setLoading(false);
        },
        (err: Error) => {
          setError(err.message);
          setItems([]);
          setLoading(false);
        }
      );
  }, [listTitle, tableConfig, effectiveSort, skip, pageSize, columnFilters, selectedViewModeId, listView, fieldMetadata]);

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
    setSkip(0);
  };

  const layout = pagination?.layout ?? 'buttons';
  const currentPage = Math.floor(skip / pageSize) + 1;
  const from = skip + 1;
  const to = skip + items.length;
  const showPagination = pagination?.enabled && (hasNext || skip > 0);
  const onPrev = (): void => setSkip(Math.max(0, skip - pageSize));
  const onNext = (): void => setSkip(skip + pageSize);
  const goToPage = (page: number): void => setSkip((page - 1) * pageSize);

  const pagedNumbers: (number | 'ellipsis')[] =
    layout === 'paged'
      ? currentPage <= 3
        ? (() => { const a: number[] = []; for (let i = 1; i <= currentPage; i++) a.push(i); return a; })()
        : [1, 'ellipsis', currentPage - 2, currentPage - 1, currentPage]
      : [];

  const btnPad = layout === 'compact' ? '4px 8px' : '6px 12px';

  const paginationBar =
    showPagination && (
      <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="end" styles={{ root: { flexWrap: 'wrap' } }}>
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
            {skip > 0 && (
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
            {skip > 0 && (
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

  const viewModes = listView?.viewModes ?? [];
  const viewModeOptions: IDropdownOption[] = viewModes.map((m) => ({ key: m.id, text: m.label }));

  if (!tableConfig) return <DataTable config={{ enabled: true, columns: [], sortable: false, emptyMessage: '' }} items={[]} loading={true} sortConfig={null} onSort={handleSort} engine={engine} />;

  return (
    <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 8 } }}>
      {viewModeOptions.length > 0 && (
        <Dropdown
          label="Visualização"
          options={viewModeOptions}
          selectedKey={selectedViewModeId}
          onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => opt && setSelectedViewModeId(String(opt.key))}
          styles={{ root: { maxWidth: 220 } }}
        />
      )}
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
      />
      {paginationBar}
    </Stack>
  );
};
