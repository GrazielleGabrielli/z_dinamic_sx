import * as React from 'react';
import type { ITableColumnConfig } from '../../core/table/types';
import type { TableEngine } from '../../core/table/services/TableEngine';
import { TableCell } from './TableCell';
import { DINAMIC_SX_TABLE_CLASS } from './tableLayoutClasses';

export interface ITableRowProps {
  item: Record<string, unknown>;
  columns: ITableColumnConfig[];
  engine: TableEngine;
}

export const TableRow: React.FC<ITableRowProps> = ({ item, columns, engine }) => (
  <tr className={DINAMIC_SX_TABLE_CLASS.row} style={{ borderBottom: '1px solid #f3f2f1' }}>
    {columns.map((col) => (
      <TableCell key={col.id} item={item} column={col} engine={engine} />
    ))}
  </tr>
);
