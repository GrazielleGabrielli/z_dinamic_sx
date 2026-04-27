import * as React from 'react';
import { useState } from 'react';
import { Callout, Stack, Text, TextField, PrimaryButton, DefaultButton } from '@fluentui/react';
import type { IListRowActionConfig } from '../../core/config/types';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import type { ITableColumnConfig, ISortConfig } from '../../core/table/types';
import type { TableEngine } from '../../core/table/services/TableEngine';
import { resolveListRowActionUrl, isSafeListRowNavigationUrl } from '../../core/table/utils/resolveListRowActionUrl';
import { TableHeader } from './TableHeader';
import { RowActionButtons } from './RowActionButtons';
import { TableEmptyState } from './TableEmptyState';
import { TableLoadingState } from './TableLoadingState';
import { TableErrorState } from './TableErrorState';
import { DINAMIC_SX_TABLE_CLASS, DINAMIC_SX_CARD_CLASS } from './tableLayoutClasses';

export interface IListItemsCardGridProps {
  columns: ITableColumnConfig[];
  items: Record<string, unknown>[];
  loading?: boolean;
  error?: string;
  emptyMessage: string;
  engine: TableEngine;
  sortConfig: ISortConfig | null;
  onSort: (field: string, direction: 'asc' | 'desc') => void;
  tableSortable: boolean;
  columnFilters?: Record<string, string>;
  onColumnFilter?: (field: string, value: string) => void;
  dense?: boolean;
  rowActions?: IListRowActionConfig[];
  dynamicContext?: IDynamicContext;
}

export const ListItemsCardGrid: React.FC<IListItemsCardGridProps> = ({
  columns,
  items,
  loading = false,
  error,
  emptyMessage,
  engine,
  sortConfig,
  onSort,
  tableSortable,
  columnFilters = {},
  onColumnFilter,
  dense,
  rowActions,
  dynamicContext,
}) => {
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
  if (columns.length === 0) return <TableEmptyState message="Nenhuma coluna visível." />;

  const body =
    loading && items.length === 0 ? (
      <TableLoadingState />
    ) : items.length === 0 ? (
      <TableEmptyState message={emptyMessage} />
    ) : (
      <div
        role="list"
        className={DINAMIC_SX_CARD_CLASS.grid}
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))',
          gap: 12,
          marginTop: 8,
        }}
      >
        {items.map((item, idx) => {
          const key = (item.Id as number | string | undefined) ?? idx;
          let wholeAction: IListRowActionConfig | undefined;
          if (rowActions) {
            for (let j = 0; j < rowActions.length; j++) {
              if (rowActions[j].scope === 'wholeRow') {
                wholeAction = rowActions[j];
                break;
              }
            }
          }
          const wholeHref =
            wholeAction !== undefined ? resolveListRowActionUrl(wholeAction.urlTemplate, item, actionContext) : '';
          const cardClickable =
            Boolean(wholeAction && wholeHref && isSafeListRowNavigationUrl(wholeHref));
          const openCard = (): void => {
            if (!cardClickable || !wholeAction || !wholeHref) return;
            if (wholeAction.openInNewTab === true) window.open(wholeHref, '_blank', 'noopener,noreferrer');
            else window.location.assign(wholeHref);
          };
          return (
            <Stack
              key={String(key)}
              role="listitem"
              className={DINAMIC_SX_CARD_CLASS.card}
              tokens={{ childrenGap: 8 }}
              onClick={cardClickable ? openCard : undefined}
              styles={{
                root: {
                  padding: 14,
                  border: '1px solid #edebe9',
                  borderRadius: 8,
                  background: '#fff',
                  boxShadow: '0 1px 2px rgba(0,0,0,0.06)',
                  minWidth: 0,
                  cursor: cardClickable ? 'pointer' : undefined,
                },
              }}
            >
              {columns.map((col, ci) => {
                const Renderer = engine.getRenderer(col);
                const resolvedValue = engine.resolveCellValue(item, col);
                const content = Renderer({ item, column: col, resolvedValue });
                if (ci === 0) {
                  return (
                    <div key={col.id} className={DINAMIC_SX_CARD_CLASS.title} style={{ fontSize: 16, fontWeight: 600, lineHeight: 1.3, wordBreak: 'break-word' }}>
                      {content}
                    </div>
                  );
                }
                return (
                  <div key={col.id} className={DINAMIC_SX_CARD_CLASS.fieldRow} style={{ display: 'flex', flexWrap: 'wrap', gap: '4px 8px', alignItems: 'baseline', fontSize: 13 }}>
                    <Text variant="small" className={DINAMIC_SX_CARD_CLASS.fieldLabel} styles={{ root: { color: '#605e5c', fontWeight: 600, flex: '0 0 auto' } }}>
                      {col.label}
                    </Text>
                    <span className={DINAMIC_SX_CARD_CLASS.fieldValue} style={{ flex: '1 1 auto', minWidth: 0, wordBreak: 'break-word' }}>{content}</span>
                  </div>
                );
              })}
              {showActionsColumn ? (
                <Stack
                  horizontal
                  verticalAlign="center"
                  horizontalAlign="end"
                  className={DINAMIC_SX_CARD_CLASS.actions}
                  tokens={{ childrenGap: 4 }}
                  styles={{
                    root: {
                      marginTop: 4,
                      paddingTop: 10,
                      borderTop: '1px solid #edebe9',
                    },
                  }}
                  onClick={(ev) => { ev.stopPropagation(); }}
                >
                  <RowActionButtons actions={rowActions ?? []} item={item} dynamicContext={actionContext} />
                </Stack>
              ) : null}
            </Stack>
          );
        })}
      </div>
    );

  return (
    <div className={DINAMIC_SX_TABLE_CLASS.scrollWrap} style={{ overflowX: 'auto' }}>
      <table
        className={DINAMIC_SX_TABLE_CLASS.table}
        role="grid"
        style={{
          width: '100%',
          borderCollapse: 'collapse',
          tableLayout: dense ? 'fixed' : 'auto',
        }}
      >
        <TableHeader
          columns={columns}
          sortConfig={sortConfig}
          onSort={onSort}
          tableSortable={tableSortable}
          columnFilters={columnFilters}
          onColumnFilter={onColumnFilter}
          onOpenFilter={handleOpenFilter}
          showActionsColumn={showActionsColumn}
        />
      </table>
      {body}
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
