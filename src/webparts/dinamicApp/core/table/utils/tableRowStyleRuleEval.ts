import type { ITableRowStyleRule, TTableRowRuleOperator } from '../../config/types';
import type { TableEngine } from '../services/TableEngine';
import type { ITableColumnConfig } from '../types';

const VALID_OPS = new Set<string>([
  'eq',
  'ne',
  'contains',
  'startsWith',
  'endsWith',
  'empty',
  'notEmpty',
]);

export function toTableRowRuleDataToken(id: string): string {
  const t = id.trim().replace(/\s+/g, '_');
  return t.replace(/[^a-zA-Z0-9_-]/g, '_') || 'rule';
}

export function sanitizeTableRowStyleRules(raw: unknown): ITableRowStyleRule[] | undefined {
  if (!Array.isArray(raw)) return undefined;
  const out: ITableRowStyleRule[] = [];
  for (let i = 0; i < raw.length; i++) {
    const x = raw[i];
    if (!x || typeof x !== 'object') continue;
    const o = x as Record<string, unknown>;
    const idRaw = typeof o.id === 'string' ? o.id.trim() : '';
    const field = typeof o.field === 'string' ? o.field.trim() : '';
    const op = typeof o.operator === 'string' ? o.operator : '';
    const value = typeof o.value === 'string' ? o.value : '';
    const rowCss = typeof o.rowCss === 'string' ? o.rowCss.trim() : '';
    if (!field || !VALID_OPS.has(op) || !rowCss) continue;
    const id = idRaw ? toTableRowRuleDataToken(idRaw) : `r_${Date.now()}_${i}`;
    out.push({
      id,
      field,
      operator: op as TTableRowRuleOperator,
      value,
      rowCss,
    });
  }
  return out.length > 0 ? out : undefined;
}

export function getRuleComparableString(
  item: Record<string, unknown>,
  field: string,
  engine: TableEngine,
  columns: ITableColumnConfig[]
): string {
  let col: ITableColumnConfig | undefined;
  for (let i = 0; i < columns.length; i++) {
    if (columns[i].internalName === field) {
      col = columns[i];
      break;
    }
  }
  if (col) {
    return String(engine.resolveDisplayValue(item, col) ?? '').trim();
  }
  const raw = item[field];
  if (raw == null || raw === '') return '';
  if (typeof raw === 'object' && raw !== null) {
    const o = raw as Record<string, unknown>;
    if (o.Title != null) return String(o.Title).trim();
    if (o.Label != null) return String(o.Label).trim();
  }
  return String(raw).trim();
}

export function evaluateTableRowStyleRule(
  item: Record<string, unknown>,
  rule: ITableRowStyleRule,
  engine: TableEngine,
  columns: ITableColumnConfig[]
): boolean {
  const cell = getRuleComparableString(item, rule.field, engine, columns);
  const val = (rule.value ?? '').trim();
  switch (rule.operator) {
    case 'eq':
      return cell === val;
    case 'ne':
      return cell !== val;
    case 'contains':
      return val === '' ? cell === '' : cell.toLowerCase().indexOf(val.toLowerCase()) !== -1;
    case 'startsWith': {
      if (val === '') return cell === '';
      const c = cell.toLowerCase();
      const v = val.toLowerCase();
      return c.length >= v.length && c.indexOf(v) === 0;
    }
    case 'endsWith': {
      if (val === '') return cell === '';
      const c = cell.toLowerCase();
      const v = val.toLowerCase();
      return c.length >= v.length && c.slice(c.length - v.length) === v;
    }
    case 'empty':
      return cell.length === 0;
    case 'notEmpty':
      return cell.length > 0;
    default:
      return false;
  }
}
