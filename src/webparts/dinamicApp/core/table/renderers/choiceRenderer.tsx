import * as React from 'react';
import type { ITableRendererProps } from '../types';

export function choiceRenderer(props: ITableRendererProps): React.ReactNode {
  const { column, resolvedValue } = props;
  if (resolvedValue == null || resolvedValue === '') return column.fallbackValue ?? '—';
  return String(resolvedValue);
}
