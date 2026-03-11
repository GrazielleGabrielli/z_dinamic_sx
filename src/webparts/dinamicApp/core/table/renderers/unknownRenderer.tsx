import * as React from 'react';
import type { ITableRendererProps } from '../types';
import { resolveDisplayCellValue, resolveFallbackValue } from '../utils/valueResolver';

export function unknownRenderer(props: ITableRendererProps): React.ReactNode {
  const { item, column, resolvedValue } = props;
  if (resolvedValue == null || resolvedValue === '') return resolveFallbackValue(column);
  const display = resolveDisplayCellValue(item, column);
  return display || resolveFallbackValue(column);
}
