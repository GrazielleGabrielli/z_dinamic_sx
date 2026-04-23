import * as React from 'react';
import type { ITableRendererProps } from '../types';
import {
  resolveDisplayCellValue,
  resolveFallbackValue,
  resolveRawCellValue,
} from '../utils/valueResolver';
import { MultiChoiceBadges, parseMultiChoiceLabels } from './multiChoiceBadges';

export function unknownRenderer(props: ITableRendererProps): React.ReactNode {
  const { item, column, resolvedValue } = props;
  if (resolvedValue == null || resolvedValue === '') return resolveFallbackValue(column);
  const raw = resolveRawCellValue(item, column);
  if (typeof raw === 'string') {
    const t = raw.trim();
    if (t.charAt(0) === '[' && t.charAt(t.length - 1) === ']') {
      const labels = parseMultiChoiceLabels(raw);
      if (labels.length > 0) {
        return <MultiChoiceBadges labels={labels} emptyFallback={resolveFallbackValue(column)} />;
      }
    }
  }
  const display = resolveDisplayCellValue(item, column);
  return display || resolveFallbackValue(column);
}
