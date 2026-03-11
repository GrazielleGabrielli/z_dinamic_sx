import * as React from 'react';
import type { ITableRendererProps } from '../types';
import { resolveExpandedValue, resolveFallbackValue } from '../utils/valueResolver';

export function userRenderer(props: ITableRendererProps): React.ReactNode {
  const { column, resolvedValue } = props;
  if (resolvedValue == null || resolvedValue === '') return resolveFallbackValue(column);
  const display = resolveExpandedValue(resolvedValue, column.expandConfig ?? { displayField: 'Title' });
  return display || resolveFallbackValue(column);
}
