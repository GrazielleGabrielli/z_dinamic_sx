import type { ITableColumnConfig } from '../types';
import { EXPANDABLE_FIELD_TYPES } from '../constants/tableDefaults';

/** Caminho OData para filtro/ordenação (ex.: `Departamento/Title`). */
export function columnODataPath(column: ITableColumnConfig): string {
  const ft = column.fieldType;
  if (ft === 'lookupMulti' || ft === 'userMulti') {
    return column.internalName;
  }
  if (ft && EXPANDABLE_FIELD_TYPES.indexOf(ft) !== -1) {
    const df = column.expandConfig?.displayField ?? 'Title';
    return `${column.internalName}/${df}`;
  }
  return column.internalName;
}
