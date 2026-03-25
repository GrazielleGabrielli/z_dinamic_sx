import * as React from 'react';
import type { ITableColumnConfig } from '../../core/table/types';
import type { TableEngine } from '../../core/table/services/TableEngine';
import { DINAMIC_SX_TABLE_CLASS } from './tableLayoutClasses';

export interface ITableCellProps {
  item: Record<string, unknown>;
  column: ITableColumnConfig;
  engine: TableEngine;
}

export const TableCell: React.FC<ITableCellProps> = ({ item, column, engine }) => {
  const resolvedValue = engine.resolveCellValue(item, column);
  const Renderer = engine.getRenderer(column);
  const content = Renderer({
    item,
    column,
    resolvedValue,
  });

  return (
    <td
      className={DINAMIC_SX_TABLE_CLASS.cell}
      data-field={column.internalName}
      style={{
        textAlign: column.align ?? 'left',
        padding: '8px 12px',
        borderBottom: '1px solid #f3f2f1',
        verticalAlign: 'middle',
      }}
    >
      {content}
    </td>
  );
};
