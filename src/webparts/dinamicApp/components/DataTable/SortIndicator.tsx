import * as React from 'react';
import { Icon } from '@fluentui/react';
import type { ISortConfig } from '../../core/table/types';

export interface ISortIndicatorProps {
  columnInternalName: string;
  sortConfig: ISortConfig | null;
}

export const SortIndicator: React.FC<ISortIndicatorProps> = ({ columnInternalName, sortConfig }) => {
  if (!sortConfig || sortConfig.field !== columnInternalName) return null;
  const icon = sortConfig.direction === 'asc' ? 'SortUp' : 'SortDown';
  return <Icon iconName={icon} styles={{ root: { marginLeft: 4, fontSize: 12 } }} />;
};
