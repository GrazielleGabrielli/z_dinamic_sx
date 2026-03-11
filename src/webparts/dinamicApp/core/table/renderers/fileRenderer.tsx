import * as React from 'react';
import type { ITableRendererProps } from '../types';

export function fileRenderer(props: ITableRendererProps): React.ReactNode {
  const { column, resolvedValue } = props;
  if (resolvedValue == null || resolvedValue === '') return column.fallbackValue ?? '—';
  const v = resolvedValue;
  if (typeof v === 'object' && v !== null) {
    const name = (v as Record<string, unknown>).Name ?? (v as Record<string, unknown>).FileName;
    const url = (v as Record<string, unknown>).ServerUrl ?? (v as Record<string, unknown>).LinkingUrl;
    if (url && name) {
      return (
        <a href={String(url)} target="_blank" rel="noopener noreferrer">
          {String(name)}
        </a>
      );
    }
  }
  return String(v);
}
