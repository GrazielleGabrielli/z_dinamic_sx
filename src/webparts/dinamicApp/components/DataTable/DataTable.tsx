import * as React from 'react';
import { useState } from 'react';
import { Callout, Stack, TextField, PrimaryButton, DefaultButton } from '@fluentui/react';
import { TableEngine } from '../../core/table/services/TableEngine';
import type { ITableConfig, ISortConfig } from '../../core/table/types';
import type { IListRowActionConfig, ITableRowStyleRule } from '../../core/config/types';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import { TableHeader } from './TableHeader';
import { TableRow } from './TableRow';
import { TableEmptyState } from './TableEmptyState';
import { TableLoadingState } from './TableLoadingState';
import { TableErrorState } from './TableErrorState';
import { DINAMIC_SX_TABLE_CLASS } from './tableLayoutClasses';

export interface IDataTableProps {
  config: ITableConfig;
  items: Record<string, unknown>[];
  loading?: boolean;
  error?: string;
  sortConfig: ISortConfig | null;
  onSort: (field: string, direction: 'asc' | 'desc') => void;
  columnFilters?: Record<string, string>;
  onColumnFilter?: (field: string, value: string) => void;
  engine: TableEngine;
  rowStyleRules?: ITableRowStyleRule[];
  rowActions?: IListRowActionConfig[];
  dynamicContext?: IDynamicContext;
  userGroupIds?: Set<number>;
}

export const DataTable: React.FC<IDataTableProps> = ({
  config,
  items,
  loading = false,
  error,
  sortConfig,
  onSort,
  columnFilters = {},
  onColumnFilter,
  engine,
  rowStyleRules,
  rowActions,
  dynamicContext,
  userGroupIds,
}) => {
  const columns = engine.getVisibleColumns(config);
  const actionContext: IDynamicContext = dynamicContext ?? { now: new Date() };
  const showActionsColumn = Boolean(rowActions && rowActions.length > 0);
  const [filterColumn, setFilterColumn] = useState<string | null>(null);
  const [filterTarget, setFilterTarget] = useState<HTMLElement | null>(null);
  const [filterInputValue, setFilterInputValue] = useState('');

  const handleOpenFilter = (field: string, target: HTMLElement): void => {
    setFilterColumn(field);
    setFilterTarget(target);
    setFilterInputValue(columnFilters[field] ?? '');
  };

  const applyFilter = (): void => {
    if (filterColumn && onColumnFilter) {
      onColumnFilter(filterColumn, filterInputValue);
      setFilterColumn(null);
      setFilterTarget(null);
    }
  };

  const clearFilter = (): void => {
    if (filterColumn && onColumnFilter) {
      onColumnFilter(filterColumn, '');
    }
    setFilterInputValue('');
    setFilterColumn(null);
    setFilterTarget(null);
  };

  if (error) return <TableErrorState message={error} />;
  if (loading && items.length === 0) return <TableLoadingState />;
  if (columns.length === 0) return <TableEmptyState message="Nenhuma coluna visível." />;
  if (items.length === 0) return <TableEmptyState message={config.emptyMessage} />;

  return (
    <div className={DINAMIC_SX_TABLE_CLASS.scrollWrap} style={{ overflowX: 'auto' }}>
      <table
        className={DINAMIC_SX_TABLE_CLASS.table}
        role="grid"
        style={{
          width: '100%',
          borderCollapse: 'collapse',
          tableLayout: config.dense ? 'fixed' : 'auto',
        }}
      >
        <TableHeader
          columns={columns}
          sortConfig={sortConfig}
          onSort={onSort}
          tableSortable={config.sortable}
          columnFilters={columnFilters}
          onColumnFilter={onColumnFilter}
          onOpenFilter={handleOpenFilter}
          showActionsColumn={showActionsColumn}
        />
        <tbody className={DINAMIC_SX_TABLE_CLASS.body}>
          {items.map((item, idx) => (
            <TableRow
              key={(item.Id as number) ?? idx}
              item={item}
              columns={columns}
              engine={engine}
              rowStyleRules={rowStyleRules}
              rowActions={rowActions}
              dynamicContext={actionContext}
              userGroupIds={userGroupIds}
            />
          ))}
        </tbody>
      </table>
      {filterColumn && filterTarget && (
        <Callout
          target={filterTarget}
          onDismiss={() => { setFilterColumn(null); setFilterTarget(null); }}
          setInitialFocus
          role="dialog"
          ariaLabel="Filtrar por campo"
        >
          <Stack tokens={{ childrenGap: 8 }} styles={{ root: { padding: 12, minWidth: 220 } }}>
            <TextField
              placeholder="Digite para filtrar..."
              value={filterInputValue}
              onChange={(_: React.FormEvent<HTMLInputElement>, v?: string) => setFilterInputValue(v ?? '')}
              onKeyDown={(ev) => ev.key === 'Enter' && applyFilter()}
            />
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton text="Limpar" onClick={clearFilter} />
              <PrimaryButton text="Filtrar" onClick={applyFilter} />
            </Stack>
          </Stack>
        </Callout>
      )}
    </div>
  );
};
