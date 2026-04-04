import * as React from 'react';
import type { ITableColumnConfig } from '../../core/table/types';
import type { TableEngine } from '../../core/table/services/TableEngine';
import type { ITableRowStyleRule } from '../../core/config/types';
import { TableCell } from './TableCell';
import { DINAMIC_SX_TABLE_CLASS } from './tableLayoutClasses';
import { evaluateTableRowStyleRule, toTableRowRuleDataToken } from '../../core/table/utils/tableRowStyleRuleEval';

export interface ITableRowProps {
  item: Record<string, unknown>;
  columns: ITableColumnConfig[];
  engine: TableEngine;
  rowStyleRules?: ITableRowStyleRule[];
}

export const TableRow: React.FC<ITableRowProps> = ({ item, columns, engine, rowStyleRules }) => {
  const matchedTokens =
    rowStyleRules
      ?.filter((r) => evaluateTableRowStyleRule(item, r, engine, columns))
      .map((r) => toTableRowRuleDataToken(r.id)) ?? [];
  const dataRules = matchedTokens.length > 0 ? matchedTokens.join(' ') : undefined;

  return (
    <tr className={DINAMIC_SX_TABLE_CLASS.row}>
      {columns.map((col) => (
        <TableCell key={col.id} item={item} column={col} engine={engine} rowDataRules={dataRules} />
      ))}
    </tr>
  );
};
