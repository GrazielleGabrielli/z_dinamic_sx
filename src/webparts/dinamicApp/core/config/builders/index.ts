import { IDynamicViewConfig, IDataSourceConfig, IDashboardConfig, IPaginationConfig, TViewMode } from '../types';
import { getDefaultConfig } from '../utils';

const DEFAULT_PAGE_SIZE_OPTIONS = [5, 10, 20, 50, 100];

export function buildConfig(params: {
  dataSource: IDataSourceConfig;
  mode: TViewMode;
  dashboard: Partial<IDashboardConfig>;
  pagination: Partial<IPaginationConfig>;
}): IDynamicViewConfig {
  const defaults = getDefaultConfig();
  return {
    dataSource: params.dataSource,
    mode: params.mode,
    dashboard: {
      enabled: params.dashboard.enabled ?? defaults.dashboard.enabled,
      cardsCount: params.dashboard.cardsCount ?? defaults.dashboard.cardsCount,
      cards: params.dashboard.cards ?? [],
    },
    pagination: {
      enabled: params.pagination.enabled ?? defaults.pagination.enabled,
      pageSize: params.pagination.pageSize ?? defaults.pagination.pageSize,
      pageSizeOptions:
        params.pagination.pageSizeOptions?.length
          ? params.pagination.pageSizeOptions
          : DEFAULT_PAGE_SIZE_OPTIONS,
    },
  };
}
