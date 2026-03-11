import type { ITableColumnConfig } from '../types';
import { EXPANDABLE_FIELD_TYPES } from '../constants/tableDefaults';

function needsExpand(column: ITableColumnConfig): boolean {
  const ft = column.fieldType;
  if (!ft) return false;
  return EXPANDABLE_FIELD_TYPES.indexOf(ft) !== -1;
}

export function buildSelect(columns: ITableColumnConfig[]): string[] {
  const select: string[] = ['Id'];
  const hasTitle = columns.some((c) => c.internalName === 'Title');

  for (const c of columns) {
    if (!c.visible || !c.internalName) continue;
    if (c.internalName === 'Id') continue;

    if (needsExpand(c)) {
      select.push(`${c.internalName}/Id`, `${c.internalName}/Title`);
    } else {
      select.push(c.internalName);
    }
  }

  if (!hasTitle && select.length === 1) select.push('Title');
  return select;
}

export function buildExpand(columns: ITableColumnConfig[]): string[] {
  const expand: string[] = [];
  for (const c of columns) {
    if (!c.visible || !c.internalName) continue;
    if (needsExpand(c) && expand.indexOf(c.internalName) === -1) {
      expand.push(c.internalName);
    }
  }
  return expand;
}

export function buildSelectExpand(columns: ITableColumnConfig[]): { select: string[]; expand: string[] } {
  return {
    select: buildSelect(columns),
    expand: buildExpand(columns),
  };
}
