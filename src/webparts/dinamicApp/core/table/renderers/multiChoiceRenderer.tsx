import * as React from 'react';
import type { ITableRendererProps } from '../types';
import { MultiChoiceBadges, parseMultiChoiceLabels } from './multiChoiceBadges';

export function multiChoiceRenderer(props: ITableRendererProps): React.ReactNode {
  const { column, resolvedValue } = props;
  const labels = parseMultiChoiceLabels(resolvedValue);
  if (labels.length === 0) return column.fallbackValue ?? '—';
  return <MultiChoiceBadges labels={labels} emptyFallback={column.fallbackValue ?? '—'} />;
}
