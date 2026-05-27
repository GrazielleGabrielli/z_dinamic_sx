import * as React from 'react';
import type { ITableColumnConfig } from '../../core/table/types';
import { columnODataPath } from '../../core/table/utils/columnODataPath';
import type { TableEngine } from '../../core/table/services/TableEngine';
import { DINAMIC_SX_TABLE_CLASS } from './tableLayoutClasses';

export interface ITableCellProps {
  item: Record<string, unknown>;
  column: ITableColumnConfig;
  engine: TableEngine;
  rowDataRules?: string;
}

export const TableCell: React.FC<ITableCellProps> = ({ item, column, engine, rowDataRules }) => {
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
      data-field={columnODataPath(column)}
      {...(rowDataRules ? { 'data-dinamic-rules': rowDataRules } : {})}
      style={{
        textAlign: column.align ?? 'left',
        width: column.width,
        minWidth: column.minWidth,
        maxWidth: column.maxWidth,
      }}
    >
      {content}
    </td>
  );
};
