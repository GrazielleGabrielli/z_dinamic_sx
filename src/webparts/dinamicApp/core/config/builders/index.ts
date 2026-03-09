import { IDynamicViewConfig, IDataSourceConfig, IDashboardConfig, IPaginationConfig, IListViewConfig, TViewMode } from '../types';
import { getDefaultConfig } from '../utils';

const DEFAULT_PAGE_SIZE_OPTIONS = [5, 10, 20, 50, 100];

export function buildConfig(params: {
  dataSource: IDataSourceConfig;
  mode: TViewMode;
  dashboard: Partial<IDashboardConfig>;
  pagination: Partial<IPaginationConfig>;
  listView?: Partial<IListViewConfig>;
}): IDynamicViewConfig {
  const defaults = getDefaultConfig();
  return {
    dataSource: params.dataSource,
    mode: params.mode,
    dashboard: {
      enabled: params.dashboard.enabled ?? defaults.dashboard.enabled,
      dashboardType: params.dashboard.dashboardType ?? defaults.dashboard.dashboardType,
      cardsCount: params.dashboard.cardsCount ?? defaults.dashboard.cardsCount,
      cards: params.dashboard.cards ?? [],
      chartType: params.dashboard.chartType ?? defaults.dashboard.chartType,
      chartSeries: params.dashboard.chartSeries ?? [],
    },
    pagination: {
      enabled: params.pagination.enabled ?? defaults.pagination.enabled,
      pageSize: params.pagination.pageSize ?? defaults.pagination.pageSize,
      pageSizeOptions:
        params.pagination.pageSizeOptions?.length
          ? params.pagination.pageSizeOptions
          : DEFAULT_PAGE_SIZE_OPTIONS,
    },
    listView: {
      columns: params.listView?.columns ?? defaults.listView.columns,
      filters: params.listView?.filters ?? defaults.listView.filters,
      sort: params.listView?.sort ?? defaults.listView.sort,
    },
  };
}
