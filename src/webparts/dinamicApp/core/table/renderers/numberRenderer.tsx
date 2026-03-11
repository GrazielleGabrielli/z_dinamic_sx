import * as React from 'react';
import type { ITableRendererProps } from '../types';
import { formatDisplayValue } from '../utils/displayValueFormatter';

export function numberRenderer(props: ITableRendererProps): React.ReactNode {
  const { column, resolvedValue } = props;
  if (resolvedValue == null || resolvedValue === '') return column.fallbackValue ?? '—';
  const n = Number(resolvedValue);
  if (isNaN(n)) return column.fallbackValue ?? '—';
  return formatDisplayValue(n, column);
}
