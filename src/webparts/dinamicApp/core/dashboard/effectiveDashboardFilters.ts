import type { IDashboardCardFilter } from '../config/types';

export function effectiveDashboardFilters(entity: {
  filter?: IDashboardCardFilter;
  filters?: IDashboardCardFilter[];
}): IDashboardCardFilter[] {
  if (entity.filters && entity.filters.length > 0) return entity.filters;
  if (entity.filter) return [entity.filter];
  return [];
}
