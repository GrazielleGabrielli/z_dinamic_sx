import * as React from 'react';
import { useState } from 'react';
import { Callout, Stack, TextField, PrimaryButton, DefaultButton } from '@fluentui/react';
import type { IListRowActionConfig } from '../../core/config/types';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import type { ITableColumnConfig, ISortConfig } from '../../core/table/types';
import type { TableEngine } from '../../core/table/services/TableEngine';
import { resolveListRowActionUrl, isSafeListRowNavigationUrl } from '../../core/table/utils/resolveListRowActionUrl';
import { checkRowActionVisibility } from '../../core/table/utils/checkRowActionVisibility';
import { ListItemsCardGridToolbar } from './ListItemsCardGridToolbar';
import { RowActionButtons } from './RowActionButtons';
import { TableEmptyState } from './TableEmptyState';
import { TableLoadingState } from './TableLoadingState';
import { TableErrorState } from './TableErrorState';
import { DINAMIC_SX_TABLE_CLASS, DINAMIC_SX_CARD_CLASS } from './tableLayoutClasses';

export interface IListItemsCardGridProps {
  columns: ITableColumnConfig[];
  displayColumns?: ITableColumnConfig[];
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
  userGroupIds?: Set<number>;
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
  rowActions,
  dynamicContext,
  userGroupIds,
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

  const showToolbar = columns.length > 0 && (tableSortable || onColumnFilter !== undefined);

  const filterCallout =
    filterColumn && filterTarget ? (
      <Callout
        target={filterTarget}
        onDismiss={() => {
          setFilterColumn(null);
          setFilterTarget(null);
        }}
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
    ) : null;

  if (columns.length === 0) {
    return (
      <div className={DINAMIC_SX_TABLE_CLASS.scrollWrap}>
        <TableEmptyState message="Nenhuma coluna visível." />
      </div>
    );
  }

  let body: React.ReactNode;
  if (loading && items.length === 0) {
    body = <TableLoadingState />;
  } else if (items.length === 0) {
    body = <TableEmptyState message={emptyMessage} />;
  } else {
    body = (
      <div role="list" className={DINAMIC_SX_CARD_CLASS.grid}>
        {items.map((item, idx) => {
          const key = (item.Id as number | string | undefined) ?? idx;
          let wholeAction: IListRowActionConfig | undefined;
          if (rowActions) {
            for (let j = 0; j < rowActions.length; j++) {
              if (
                rowActions[j].scope === 'wholeRow' &&
                checkRowActionVisibility(rowActions[j], item, actionContext, userGroupIds)
              ) {
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
          const titleCol = columns[0];
          const detailCols = columns.slice(1);
          const TitleRenderer = titleCol ? engine.getRenderer(titleCol) : null;
          const titleContent =
            titleCol && TitleRenderer
              ? TitleRenderer({
                  item,
                  column: titleCol,
                  resolvedValue: engine.resolveCellValue(item, titleCol),
                })
              : null;

          return (
            <article
              key={String(key)}
              role="listitem"
              className={`${DINAMIC_SX_CARD_CLASS.card}${cardClickable ? ` ${DINAMIC_SX_CARD_CLASS.cardClickable}` : ''}`}
              onClick={cardClickable ? openCard : undefined}
              onKeyDown={
                cardClickable
                  ? (ev) => {
                      if (ev.key === 'Enter' || ev.key === ' ') {
                        ev.preventDefault();
                        openCard();
                      }
                    }
                  : undefined
              }
              tabIndex={cardClickable ? 0 : undefined}
            >
              {titleCol ? (
                <header className={DINAMIC_SX_CARD_CLASS.cardHeader}>
                  <div className={DINAMIC_SX_CARD_CLASS.title}>{titleContent}</div>
                </header>
              ) : null}
              {detailCols.length > 0 ? (
                <div className={DINAMIC_SX_CARD_CLASS.cardBody}>
                  {detailCols.map((col) => {
                    const Renderer = engine.getRenderer(col);
                    const resolvedValue = engine.resolveCellValue(item, col);
                    const content = Renderer({ item, column: col, resolvedValue });
                    return (
                      <div key={col.id} className={DINAMIC_SX_CARD_CLASS.fieldRow}>
                        <span className={DINAMIC_SX_CARD_CLASS.fieldLabel}>{col.label}</span>
                        <span className={DINAMIC_SX_CARD_CLASS.fieldValue}>{content}</span>
                      </div>
                    );
                  })}
                </div>
              ) : null}
              {showActionsColumn ? (
                <footer
                  className={DINAMIC_SX_CARD_CLASS.actions}
                  onClick={(ev) => {
                    ev.stopPropagation();
                  }}
                >
                  <RowActionButtons
                    actions={rowActions ?? []}
                    item={item}
                    dynamicContext={actionContext}
                    userGroupIds={userGroupIds}
                  />
                </footer>
              ) : null}
            </article>
          );
        })}
      </div>
    );
  }

  return (
    <>
      <div className={DINAMIC_SX_TABLE_CLASS.scrollWrap}>
        {showToolbar ? (
          <ListItemsCardGridToolbar
            columns={columns}
            sortConfig={sortConfig}
            onSort={onSort}
            tableSortable={tableSortable}
            columnFilters={columnFilters}
            onOpenFilter={onColumnFilter ? handleOpenFilter : undefined}
            showActionsColumn={showActionsColumn}
          />
        ) : null}
        {body}
      </div>
      {filterCallout}
    </>
  );
};
