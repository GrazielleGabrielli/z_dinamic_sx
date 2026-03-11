import * as React from 'react';
import type { ITableColumnConfig } from '../../core/table/types';
import type { TableEngine } from '../../core/table/services/TableEngine';
import { TableCell } from './TableCell';

export interface ITableRowProps {
  item: Record<string, unknown>;
  columns: ITableColumnConfig[];
  engine: TableEngine;
}

export const TableRow: React.FC<ITableRowProps> = ({ item, columns, engine }) => (
  <tr style={{ borderBottom: '1px solid #f3f2f1' }}>
    {columns.map((col) => (
      <TableCell key={col.id} item={item} column={col} engine={engine} />
    ))}
  </tr>
);
