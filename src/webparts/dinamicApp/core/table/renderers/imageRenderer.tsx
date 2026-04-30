import * as React from 'react';
import { ensureAbsoluteSharePointUrl } from '../../formManager/formUrlUtils';
import type { ITableRendererProps } from '../types';

export function imageRenderer(props: ITableRendererProps): React.ReactNode {
  const { column, resolvedValue } = props;
  if (resolvedValue == null || resolvedValue === '') return column.fallbackValue ?? '—';
  if (typeof resolvedValue === 'object' && resolvedValue !== null && 'Url' in resolvedValue) {
    const url = ensureAbsoluteSharePointUrl(String((resolvedValue as { Url: string }).Url ?? ''));
    if (url) return <img src={url} alt="" style={{ maxWidth: 48, maxHeight: 48 }} />;
  }
  if (typeof resolvedValue === 'string' && resolvedValue.trim()) {
    return <img src={ensureAbsoluteSharePointUrl(resolvedValue)} alt="" style={{ maxWidth: 48, maxHeight: 48 }} />;
  }
  return String(resolvedValue);
}
