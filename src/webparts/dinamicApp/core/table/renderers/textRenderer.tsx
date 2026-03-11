import * as React from 'react';
import type { ITableRendererProps } from '../types';

export function textRenderer(props: ITableRendererProps): React.ReactNode {
  const { column, resolvedValue } = props;
  const raw = resolvedValue;
  if (raw == null || raw === '') return column.fallbackValue ?? '—';
  return String(raw);
}
