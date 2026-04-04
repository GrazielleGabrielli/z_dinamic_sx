import * as React from 'react';
import type { ITableColumnConfig } from '../../core/table/types';
import type { TableEngine } from '../../core/table/services/TableEngine';
import type { IListRowActionConfig, ITableRowStyleRule } from '../../core/config/types';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import { TableCell } from './TableCell';
import { RowActionButtons } from './RowActionButtons';
import { DINAMIC_SX_TABLE_CLASS } from './tableLayoutClasses';
import { evaluateTableRowStyleRule, toTableRowRuleDataToken } from '../../core/table/utils/tableRowStyleRuleEval';
import { resolveListRowActionUrl, isSafeListRowNavigationUrl } from '../../core/table/utils/resolveListRowActionUrl';

export interface ITableRowProps {
  item: Record<string, unknown>;
  columns: ITableColumnConfig[];
  engine: TableEngine;
  rowStyleRules?: ITableRowStyleRule[];
  rowActions?: IListRowActionConfig[];
  dynamicContext: IDynamicContext;
}

export const TableRow: React.FC<ITableRowProps> = ({ item, columns, engine, rowStyleRules, rowActions, dynamicContext }) => {
  const matchedTokens =
    rowStyleRules
      ?.filter((r) => evaluateTableRowStyleRule(item, r, engine, columns))
      .map((r) => toTableRowRuleDataToken(r.id)) ?? [];
  const dataRules = matchedTokens.length > 0 ? matchedTokens.join(' ') : undefined;

  let wholeAction: IListRowActionConfig | undefined;
  if (rowActions) {
    for (let i = 0; i < rowActions.length; i++) {
      if (rowActions[i].scope === 'wholeRow') {
        wholeAction = rowActions[i];
        break;
      }
    }
  }
  const wholeHref =
    wholeAction !== undefined ? resolveListRowActionUrl(wholeAction.urlTemplate, item, dynamicContext) : '';
  const rowClickable =
    Boolean(wholeAction && wholeHref && isSafeListRowNavigationUrl(wholeHref));

  const openWholeRow = (): void => {
    if (!rowClickable || !wholeAction || !wholeHref) return;
    if (wholeAction.openInNewTab === true) window.open(wholeHref, '_blank', 'noopener,noreferrer');
    else window.location.assign(wholeHref);
  };

  return (
    <tr
      className={DINAMIC_SX_TABLE_CLASS.row}
      onClick={rowClickable ? openWholeRow : undefined}
      style={rowClickable ? { cursor: 'pointer' } : undefined}
    >
      {columns.map((col) => (
        <TableCell key={col.id} item={item} column={col} engine={engine} rowDataRules={dataRules} />
      ))}
      {rowActions && rowActions.length > 0 ? (
        <td
          className={DINAMIC_SX_TABLE_CLASS.cell}
          data-field="__actions"
          onClick={(ev) => { ev.stopPropagation(); }}
          style={{
            textAlign: 'right',
            padding: '8px 12px',
            borderBottom: '1px solid #f3f2f1',
            verticalAlign: 'middle',
            width: 1,
            whiteSpace: 'nowrap',
          }}
        >
          <RowActionButtons actions={rowActions} item={item} dynamicContext={dynamicContext} />
        </td>
      ) : null}
    </tr>
  );
};
