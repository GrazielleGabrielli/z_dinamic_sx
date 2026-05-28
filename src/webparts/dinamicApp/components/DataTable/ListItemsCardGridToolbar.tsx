import * as React from 'react';
import { IconButton } from '@fluentui/react';
import type { ITableColumnConfig, ISortConfig } from '../../core/table/types';
import { columnODataPath } from '../../core/table/utils/columnODataPath';
import { toggleSortDirection } from '../../core/table/utils/sortBuilder';
import { DINAMIC_SX_CARD_CLASS } from './tableLayoutClasses';

export interface IListItemsCardGridToolbarProps {
  columns: ITableColumnConfig[];
  sortConfig: ISortConfig | null;
  onSort: (field: string, direction: 'asc' | 'desc') => void;
  tableSortable: boolean;
  columnFilters?: Record<string, string>;
  onOpenFilter?: (field: string, target: HTMLElement) => void;
  showActionsColumn?: boolean;
  actionsColumnLabel?: string;
}

export const ListItemsCardGridToolbar: React.FC<IListItemsCardGridToolbarProps> = ({
  columns,
  sortConfig,
  onSort,
  tableSortable,
  columnFilters = {},
  onOpenFilter,
  showActionsColumn,
  actionsColumnLabel,
}) => {
  const showFilterSort = tableSortable && onOpenFilter !== undefined;

  const handleSortClick = (col: ITableColumnConfig, ev: React.MouseEvent<unknown>): void => {
    ev.stopPropagation();
    if (!tableSortable || !col.sortable) return;
    const oPath = columnODataPath(col);
    const nextDir = sortConfig?.field === oPath ? toggleSortDirection(sortConfig.direction) : 'asc';
    onSort(oPath, nextDir);
  };

  return (
    <div className={DINAMIC_SX_CARD_CLASS.toolbar} role="toolbar" aria-label="Ordenação e filtros por coluna">
      <div className={DINAMIC_SX_CARD_CLASS.toolbarInner}>
        {columns.map((col) => {
          const oPath = columnODataPath(col);
          const isSorted = sortConfig?.field === oPath;
          const hasFilter = Boolean((columnFilters[oPath] ?? '').trim());
          return (
            <div key={col.id} className={DINAMIC_SX_CARD_CLASS.toolbarItem}>
              <span className={DINAMIC_SX_CARD_CLASS.toolbarItemLabel}>{col.label}</span>
              {showFilterSort ? (
                <>
                  <span
                    className={DINAMIC_SX_CARD_CLASS.toolbarFilterTrigger}
                    role="presentation"
                    onClick={(ev) => {
                      ev.stopPropagation();
                      onOpenFilter(oPath, ev.currentTarget as HTMLElement);
                    }}
                    style={{ display: 'inline-flex' }}
                  >
                    <IconButton
                      iconProps={{ iconName: hasFilter ? 'FilterSolid' : 'Filter' }}
                      title="Filtrar"
                      ariaLabel={`Filtrar por ${col.label}`}
                      styles={{
                        root: { width: 28, height: 28 },
                        icon: { fontSize: 12, color: hasFilter ? '#0f6cbd' : undefined },
                      }}
                    />
                  </span>
                  {col.sortable ? (
                    <IconButton
                      iconProps={{
                        iconName: isSorted
                          ? sortConfig!.direction === 'asc'
                            ? 'SortUp'
                            : 'SortDown'
                          : 'Sort',
                      }}
                      title={
                        isSorted
                          ? sortConfig!.direction === 'asc'
                            ? 'Ordenação ascendente'
                            : 'Ordenação descendente'
                          : 'Ordenar'
                      }
                      ariaLabel={`Ordenar por ${col.label}`}
                      onClick={(ev) => handleSortClick(col, ev)}
                      styles={{
                        root: { width: 28, height: 28 },
                        icon: { fontSize: 12, color: isSorted ? '#0f6cbd' : undefined },
                      }}
                    />
                  ) : null}
                </>
              ) : null}
            </div>
          );
        })}
        {showActionsColumn ? (
          <div className={DINAMIC_SX_CARD_CLASS.toolbarItem}>
            <span className={DINAMIC_SX_CARD_CLASS.toolbarItemLabel}>{actionsColumnLabel ?? 'Ações'}</span>
          </div>
        ) : null}
      </div>
    </div>
  );
};
