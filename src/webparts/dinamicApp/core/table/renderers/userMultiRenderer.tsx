import * as React from 'react';
import type { ITableRendererProps } from '../types';
import { resolveExpandedValue, resolveFallbackValue } from '../utils/valueResolver';

export function userMultiRenderer(props: ITableRendererProps): React.ReactNode {
  const { column, resolvedValue } = props;
  if (resolvedValue == null || (Array.isArray(resolvedValue) && resolvedValue.length === 0)) {
    return resolveFallbackValue(column);
  }
  const display = resolveExpandedValue(resolvedValue, column.expandConfig ?? { displayField: 'Title', separator: ', ' });
  return display || resolveFallbackValue(column);
}
