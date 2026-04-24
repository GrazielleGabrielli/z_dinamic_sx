import type { IFieldMetadata } from '../../../../services/shared/types';
import type { IListPageAlertCountRule } from '../config/types';

export type TAlertCountFilterFieldOp = 'eq' | 'ne' | 'gt' | 'ge' | 'lt' | 'le' | 'contains';

const FIELD_OPS: TAlertCountFilterFieldOp[] = ['eq', 'ne', 'gt', 'ge', 'lt', 'le', 'contains'];

function normalizeFieldOp(raw: string | undefined): TAlertCountFilterFieldOp {
  const s = (raw ?? 'eq').toLowerCase();
  return FIELD_OPS.indexOf(s as TAlertCountFilterFieldOp) !== -1 ? (s as TAlertCountFilterFieldOp) : 'eq';
}

function escapeODataString(s: string): string {
  return `'${String(s).replace(/'/g, "''")}'`;
}

export function operatorsForCountFilterField(meta: IFieldMetadata | undefined): TAlertCountFilterFieldOp[] {
  if (!meta) return ['eq', 'ne', 'contains'];
  switch (meta.MappedType) {
    case 'number':
    case 'currency':
      return ['eq', 'ne', 'gt', 'ge', 'lt', 'le'];
    case 'boolean':
      return ['eq'];
    case 'datetime':
      return ['eq', 'ne', 'ge', 'le', 'gt', 'lt'];
    case 'choice':
      return ['eq', 'ne'];
    case 'multichoice':
      return ['contains', 'eq'];
    case 'lookup':
    case 'user':
      return ['eq', 'ne'];
    default:
      return ['eq', 'ne', 'contains'];
  }
}

export function buildODataFilterFromStructured(
  meta: IFieldMetadata | undefined,
  fieldInternal: string,
  opRaw: string | undefined,
  rawValue: string | undefined
): string | undefined {
  const f = fieldInternal?.trim();
  if (!f) return undefined;
  const op = normalizeFieldOp(opRaw);
  const val = (rawValue ?? '').trim();
  const mt = meta?.MappedType;

  if (mt === 'boolean') {
    if (!val) return undefined;
    const n = val === '1' || val === 'true' || val.toLowerCase() === 'sim' ? 1 : 0;
    return `${f} eq ${n}`;
  }

  if (mt === 'number' || mt === 'currency') {
    if (val === '') return undefined;
    const n = Number(val);
    if (Number.isNaN(n)) return undefined;
    if (['gt', 'lt', 'ge', 'le', 'eq', 'ne'].indexOf(op) === -1) return `${f} eq ${n}`;
    return `${f} ${op} ${n}`;
  }

  if (mt === 'datetime') {
    if (!val) return undefined;
    if (val.startsWith("datetime'")) return `${f} ${op} ${val}`;
    return `${f} ${op} datetime'${val.replace(/'/g, "''")}'`;
  }

  if (mt === 'lookup' || mt === 'user') {
    if (val === '') return undefined;
    const id = parseInt(val, 10);
    if (Number.isNaN(id)) return undefined;
    const idField = `${f}Id`;
    if (op !== 'eq' && op !== 'ne') return `${idField} eq ${id}`;
    return `${idField} ${op} ${id}`;
  }

  if (mt === 'multichoice') {
    if (!val) return undefined;
    if (op === 'contains') return `substringof(${escapeODataString(val)}, ${f})`;
    return `${f} ${op} ${escapeODataString(val)}`;
  }

  if (mt === 'choice' || mt === 'text' || mt === 'multiline' || mt === 'url' || mt === 'calculated' || !mt) {
    if (!val) return undefined;
    if (op === 'contains') return `substringof(${escapeODataString(val)}, ${f})`;
    return `${f} ${op} ${escapeODataString(val)}`;
  }

  if (!val) return undefined;
  if (op === 'contains') return `substringof(${escapeODataString(val)}, ${f})`;
  return `${f} ${op} ${escapeODataString(val)}`;
}

export function mergeCountRuleODataFromStructured(
  rule: IListPageAlertCountRule,
  fieldsByInternal: Map<string, IFieldMetadata>
): IListPageAlertCountRule {
  if (rule.countFilterUseManualOdata) {
    return { ...rule };
  }
  const field = rule.countFilterField?.trim();
  if (!field) {
    return { ...rule };
  }
  const meta = fieldsByInternal.get(field);
  const allowed = operatorsForCountFilterField(meta);
  let op = normalizeFieldOp(rule.countFilterFieldOp);
  if (allowed.indexOf(op) === -1) op = allowed[0];
  const built = buildODataFilterFromStructured(meta, field, op, rule.countFilterValue);
  return {
    ...rule,
    countFilterField: field,
    countFilterFieldOp: op,
    countFilterValue: rule.countFilterValue,
    countFilterUseManualOdata: false,
    ...(built ? { odataFilter: built } : { odataFilter: undefined }),
  };
}

export function isFieldEligibleForAlertCountFilter(meta: IFieldMetadata): boolean {
  if (meta.Hidden) return false;
  if (meta.MappedType === 'multiline' && meta.RichText) return false;
  if (meta.MappedType === 'taxonomymulti') return false;
  return true;
}
