import type { IFormCompareRef, TFormConditionNode, TFormConditionOp } from '../config/types/formManager';

function sanitizeCompareRef(raw: unknown): IFormCompareRef | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const r = raw as Record<string, unknown>;
  const kind = r.kind === 'field' || r.kind === 'token' || r.kind === 'literal' ? r.kind : 'literal';
  const value = typeof r.value === 'string' ? r.value : String(r.value ?? '');
  return { kind, value };
}

export function sanitizeConditionNode(raw: unknown): TFormConditionNode | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const n = raw as Record<string, unknown>;
  if (n.kind === 'all' || n.kind === 'any') {
    const childrenRaw = Array.isArray(n.children) ? n.children : [];
    const children: TFormConditionNode[] = [];
    for (let i = 0; i < childrenRaw.length; i++) {
      const c = sanitizeConditionNode(childrenRaw[i]);
      if (c) children.push(c);
    }
    if (children.length === 0) return undefined;
    return { kind: n.kind, children };
  }
  const leafLike =
    n.kind === 'leaf' ||
    (typeof n.field === 'string' &&
      n.field.trim() &&
      typeof n.op === 'string' &&
      n.kind !== 'all' &&
      n.kind !== 'any');
  if (leafLike) {
    const field = typeof n.field === 'string' ? n.field.trim() : '';
    const opRaw = typeof n.op === 'string' ? n.op : 'eq';
    const ops = new Set<string>([
      'eq',
      'ne',
      'gt',
      'ge',
      'lt',
      'le',
      'contains',
      'startsWith',
      'endsWith',
      'isEmpty',
      'isFilled',
      'isTrue',
      'isFalse',
    ]);
    const op: TFormConditionOp = ops.has(opRaw) ? (opRaw as TFormConditionOp) : 'eq';
    if (!field) return undefined;
    const compare = sanitizeCompareRef(n.compare);
    return { kind: 'leaf', field, op: op as never, ...(compare ? { compare } : {}) };
  }
  return undefined;
}
