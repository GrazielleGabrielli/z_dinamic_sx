import * as React from 'react';
import { Icon, IconButton } from '@fluentui/react';
import type { ITableColumnConfig, ISortConfig } from '../../core/table/types';
import { toggleSortDirection } from '../../core/table/utils/sortBuilder';
import { DINAMIC_SX_TABLE_CLASS } from './tableLayoutClasses';

export interface ITableHeaderProps {
  columns: ITableColumnConfig[];
  sortConfig: ISortConfig | null;
  onSort: (field: string, direction: 'asc' | 'desc') => void;
  tableSortable: boolean;
  columnFilters?: Record<string, string>;
  onColumnFilter?: (field: string, value: string) => void;
  onOpenFilter?: (field: string, target: HTMLElement) => void;
  showActionsColumn?: boolean;
  actionsColumnLabel?: string;
}

export const TableHeader: React.FC<ITableHeaderProps> = ({
  columns,
  sortConfig,
  onSort,
  tableSortable,
  onColumnFilter,
  onOpenFilter,
  showActionsColumn,
  actionsColumnLabel,
}) => {
  const handleSortClick = (col: ITableColumnConfig, ev: React.MouseEvent<unknown>): void => {
    ev.stopPropagation();
    if (!tableSortable || !col.sortable) return;
    const nextDir = sortConfig?.field === col.internalName
      ? toggleSortDirection(sortConfig.direction)
      : 'asc';
    onSort(col.internalName, nextDir);
  };

  const showFilterSort = tableSortable && onColumnFilter && onOpenFilter;

  return (
    <thead className={DINAMIC_SX_TABLE_CLASS.thead}>
      <tr className={DINAMIC_SX_TABLE_CLASS.headerRow}>
        {columns.map((col) => (
          <th
            key={col.id}
            className={DINAMIC_SX_TABLE_CLASS.headerCell}
            data-field={col.internalName}
            style={{
              textAlign: col.align ?? 'left',
              minWidth: col.minWidth,
              maxWidth: col.maxWidth,
              width: col.width,
              padding: '8px 12px',
              borderBottom: '1px solid #edebe9',
              fontWeight: 600,
            }}
          >
            <span className={DINAMIC_SX_TABLE_CLASS.headerCellInner} style={{ display: 'inline-flex', alignItems: 'center', gap: 2 }}>
              {col.label}
              {showFilterSort && col.sortable && (
                <>
                  <span
                    className={DINAMIC_SX_TABLE_CLASS.headerFilterTrigger}
                    role="presentation"
                    onClick={(ev) => { ev.stopPropagation(); onOpenFilter(col.internalName, ev.currentTarget as HTMLElement); }}
                    style={{ display: 'inline-flex', cursor: 'pointer' }}
                  >
                    <IconButton
                      iconProps={{ iconName: 'Filter' }}
                      title="Filtrar"
                      ariaLabel="Filtrar por este campo"
                      styles={{ root: { width: 24, height: 24 }, icon: { fontSize: 12 } }}
                    />
                  </span>
                  <IconButton
                    iconProps={{
                      iconName: sortConfig?.field === col.internalName
                        ? (sortConfig.direction === 'asc' ? 'SortUp' : 'SortDown')
                        : 'Sort',
                    }}
                    title={sortConfig?.field === col.internalName ? (sortConfig.direction === 'asc' ? 'Ordenação ascendente' : 'Ordenação descendente') : 'Ordenar'}
                    ariaLabel="Ordenar"
                    onClick={(ev) => handleSortClick(col, ev)}
                    styles={{ root: { width: 24, height: 24 }, icon: { fontSize: 12 } }}
                  />
                </>
              )}
              {!showFilterSort && tableSortable && col.sortable && (
                <Icon
                  iconName={sortConfig?.field === col.internalName ? (sortConfig.direction === 'asc' ? 'SortUp' : 'SortDown') : 'Sort'}
                  styles={{ root: { marginLeft: 4, fontSize: 12, cursor: 'pointer' } }}
                  onClick={(ev) => handleSortClick(col, ev)}
                />
              )}
            </span>
          </th>
        ))}
        {showActionsColumn ? (
          <th
            className={DINAMIC_SX_TABLE_CLASS.headerCell}
            data-field="__actions"
            style={{
              textAlign: 'right',
              width: 1,
              padding: '8px 12px',
              borderBottom: '1px solid #edebe9',
              fontWeight: 600,
              whiteSpace: 'nowrap',
            }}
          >
            {actionsColumnLabel ?? 'Ações'}
          </th>
        ) : null}
      </tr>
    </thead>
  );
};
