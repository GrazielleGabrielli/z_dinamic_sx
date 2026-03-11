import * as React from 'react';
import type { ITableRendererProps } from '../types';

export function noteRenderer(props: ITableRendererProps): React.ReactNode {
  const { column, resolvedValue } = props;
  if (resolvedValue == null || resolvedValue === '') return column.fallbackValue ?? '—';
  const str = String(resolvedValue);
  return str.length > 100 ? str.slice(0, 100) + '…' : str;
}
