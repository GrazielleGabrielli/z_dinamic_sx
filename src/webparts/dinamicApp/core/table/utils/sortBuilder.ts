import type { ISortConfig } from '../types';

export function buildOrderBy(sortConfig: ISortConfig | null | undefined): { field: string; ascending: boolean } | undefined {
  if (!sortConfig || !sortConfig.field) return undefined;
  return {
    field: sortConfig.field,
    ascending: sortConfig.direction === 'asc',
  };
}

export function toggleSortDirection(current: 'asc' | 'desc'): 'asc' | 'desc' {
  return current === 'asc' ? 'desc' : 'asc';
}
