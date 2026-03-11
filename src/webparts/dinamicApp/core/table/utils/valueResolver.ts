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
    const v = obj[f];
    if (v != null && v !== '') parts.push(String(v));
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
