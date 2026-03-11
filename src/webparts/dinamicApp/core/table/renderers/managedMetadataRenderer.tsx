import * as React from 'react';
import type { ITableRendererProps } from '../types';

function getLabel(v: unknown): string {
  if (v == null) return '';
  if (typeof v === 'object' && v !== null && 'Label' in v) return String((v as { Label: string }).Label);
  if (typeof v === 'object' && v !== null && 'Term' in v) {
    const t = (v as { Term: { Label?: string } }).Term;
    return t?.Label ?? '';
  }
  return String(v);
}

export function managedMetadataRenderer(props: ITableRendererProps): React.ReactNode {
  const { column, resolvedValue } = props;
  if (resolvedValue == null || resolvedValue === '') return column.fallbackValue ?? '—';
  if (Array.isArray(resolvedValue)) {
    return resolvedValue.map(getLabel).filter(Boolean).join(', ');
  }
  return getLabel(resolvedValue) || (column.fallbackValue ?? '—');
}
