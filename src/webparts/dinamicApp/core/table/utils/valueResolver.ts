import type {
  ITableColumnConfig,
  ITableColumnExpandConfig,
  TCellResolvedValue,
} from '../types';
import { DEFAULT_FALLBACK_VALUE, DEFAULT_SEPARATOR_MULTI } from '../constants/tableDefaults';

function getByPath(obj: unknown, path: string): unknown {
  if (obj == null) return undefined;
  const parts = path.split('.');
  let current: unknown = obj;
  for (let i = 0; i < parts.length; i++) {
    if (current == null || typeof current !== 'object') return undefined;
    current = (current as Record<string, unknown>)[parts[i]];
  }
  return current;
}

/** Caminho com `/` (ex.: `Gerente/Title` ou `Membros/Title` em coleção). */
function resolveExpandedPath(value: unknown, path: string): string {
  const p = path.trim();
  if (!p) return '';
  if (p.indexOf('/') === -1) {
    if (value == null) return '';
    if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
      const v = (value as Record<string, unknown>)[p];
      return v != null && v !== '' ? String(v) : '';
    }
    return String(value ?? '');
  }
  const [head, ...tails] = p.split('/');
  const tail = tails.join('/');
  if (value == null || typeof value !== 'object') return '';
  const child = (value as Record<string, unknown>)[head];
  if (Array.isArray(child)) {
    return child
      .map((c) => resolveExpandedPath(c, tail))
      .filter((s) => s.length > 0)
      .join(', ');
  }
  return resolveExpandedPath(child, tail);
}

export function resolveRawCellValue(
  item: Record<string, unknown>,
  column: ITableColumnConfig
): TCellResolvedValue {
  const name = column.internalName;
  const source = column.source;
  if (source?.kind === 'externalListField') {
    return undefined;
  }
  const path = column.expandConfig?.nestedPath ?? name;
  const val = path.indexOf('.') !== -1 ? getByPath(item, path) : item[name];
  return val as TCellResolvedValue;
}

function resolveSingleExpandedValue(
  value: unknown,
  expandConfig: ITableColumnExpandConfig
): string {
  if (value == null) return '';
  if (typeof value !== 'object') return String(value);
  const obj = value as Record<string, unknown>;
  const displayFields = expandConfig.displayFields ?? (expandConfig.displayField ? [expandConfig.displayField] : ['Title']);
  const parts: string[] = [];
  for (const f of displayFields) {
    let s = '';
    if (f.indexOf('/') !== -1) {
      s = resolveExpandedPath(obj, f);
    } else {
      const v = obj[f];
      s = v != null && v !== '' ? String(v) : '';
    }
    if (s) parts.push(s);
  }
  return parts.join(' - ');
}

function resolveCollectionValue(
  values: unknown,
  expandConfig: ITableColumnExpandConfig
): string {
  if (!Array.isArray(values)) return resolveSingleExpandedValue(values, expandConfig);
  const sep = expandConfig.separator ?? DEFAULT_SEPARATOR_MULTI;
  return values
    .map((v) => resolveSingleExpandedValue(v, expandConfig))
    .filter(Boolean)
    .join(sep);
}

/** Uma etiqueta por item (lookup multi / coleção expandida), para badges na tabela. */
export function expandCollectionToLabels(
  value: unknown,
  expandConfig: ITableColumnExpandConfig
): string[] {
  if (value == null || value === '') return [];
  if (Array.isArray(value)) {
    return value
      .map((v) => resolveSingleExpandedValue(v, expandConfig))
      .filter((s) => s.length > 0);
  }
  const one = resolveSingleExpandedValue(value, expandConfig);
  return one ? [one] : [];
}

export function resolveExpandedValue(
  value: TCellResolvedValue,
  expandConfig: ITableColumnExpandConfig | undefined
): string {
  if (value == null || value === '') return '';
  if (!expandConfig) {
    if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
      const obj = value as Record<string, unknown>;
      return (obj.Title as string) ?? String(value);
    }
    return String(value);
  }
  if (Array.isArray(value)) return resolveCollectionValue(value, expandConfig);
  if (typeof value === 'object' && value !== null) {
    return resolveSingleExpandedValue(value, expandConfig);
  }
  return String(value);
}

export function resolveDisplayCellValue(
  item: Record<string, unknown>,
  column: ITableColumnConfig
): string {
  const raw = resolveRawCellValue(item, column);
  if (raw == null || raw === '') return resolveFallbackValue(column);
  const expandConfig = column.expandConfig;
  if (expandConfig || (typeof raw === 'object' && raw !== null && !Array.isArray(raw))) {
    return resolveExpandedValue(raw, expandConfig ?? { displayField: 'Title' });
  }
  if (Array.isArray(raw)) {
    return resolveExpandedValue(raw, expandConfig ?? { displayField: 'Title', separator: DEFAULT_SEPARATOR_MULTI });
  }
  return String(raw);
}

export function resolveFallbackValue(column: ITableColumnConfig): string {
  return column.fallbackValue ?? column.formatting?.emptyValueText ?? DEFAULT_FALLBACK_VALUE;
}
