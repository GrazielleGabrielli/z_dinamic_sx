import * as React from 'react';
import type { ITableRendererProps } from '../types';

export function multiChoiceRenderer(props: ITableRendererProps): React.ReactNode {
  const { column, resolvedValue } = props;
  if (resolvedValue == null || resolvedValue === '') return column.fallbackValue ?? '—';
  if (Array.isArray(resolvedValue)) return resolvedValue.join(', ');
  const str = String(resolvedValue);
  return str.indexOf(';') !== -1 ? str.replace(/;/g, ', ') : str;
}
