import * as React from 'react';
import type { ITableRendererProps } from '../types';

export function urlRenderer(props: ITableRendererProps): React.ReactNode {
  const { column, resolvedValue } = props;
  if (resolvedValue == null || resolvedValue === '') return column.fallbackValue ?? '—';
  const v = resolvedValue;
  if (typeof v === 'object' && v !== null && 'Url' in v) {
    const url = (v as { Url?: string }).Url;
    const desc = (v as { Description?: string }).Description;
    if (url) {
      return (
        <a href={url} target="_blank" rel="noopener noreferrer">
          {desc || url}
        </a>
      );
    }
  }
  return String(v);
}
