import * as React from 'react';
import type { ITableRendererProps } from '../types';
import { expandCollectionToLabels, resolveFallbackValue } from '../utils/valueResolver';
import { MultiChoiceBadges } from './multiChoiceBadges';

export function lookupMultiRenderer(props: ITableRendererProps): React.ReactNode {
  const { column, resolvedValue } = props;
  const expandConfig = column.expandConfig ?? { displayField: 'Title' };

  let labels = expandCollectionToLabels(resolvedValue, expandConfig);

  if (
    labels.length === 0 &&
    typeof resolvedValue === 'string' &&
    resolvedValue.trim().charAt(0) === '['
  ) {
    try {
      const parsed = JSON.parse(resolvedValue.trim()) as unknown;
      labels = expandCollectionToLabels(parsed, expandConfig);
    } catch {
      labels = [];
    }
  }

  if (labels.length === 0) {
    return resolveFallbackValue(column);
  }
  return <MultiChoiceBadges labels={labels} emptyFallback={resolveFallbackValue(column)} />;
}
