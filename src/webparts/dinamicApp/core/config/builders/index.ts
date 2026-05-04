import {
  IDynamicViewConfig,
  IDataSourceConfig,
  IDashboardConfig,
  IPaginationConfig,
  IListViewConfig,
  IListPageLayoutConfig,
  IFormManagerConfig,
  TViewMode,
} from '../types';
import { getDefaultConfig } from '../utils';

const DEFAULT_PAGE_SIZE_OPTIONS = [5, 10, 20, 50, 100];

export function buildConfig(params: {
  dataSource: IDataSourceConfig;
  mode: TViewMode;
  dashboard: Partial<IDashboardConfig>;
  pagination: Partial<IPaginationConfig>;
  listView?: Partial<IListViewConfig>;
  projectManagement?: IDynamicViewConfig['projectManagement'];
  listPageLayout?: IListPageLayoutConfig;
  formManager?: IFormManagerConfig;
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
      viewModes: params.listView?.viewModes ?? defaults.listView.viewModes,
      activeViewModeId: params.listView?.activeViewModeId ?? defaults.listView.activeViewModeId,
      pdfExportEnabled: params.listView?.pdfExportEnabled ?? defaults.listView.pdfExportEnabled,
      listCardViewEnabled: params.listView?.listCardViewEnabled ?? defaults.listView.listCardViewEnabled ?? false,
      ...(params.listView?.listCardViewEnabled === true && params.listView?.listDefaultDisplayMode === 'cards'
        ? { listDefaultDisplayMode: 'cards' as const }
        : {}),
      customTableCssSlots: params.listView?.customTableCssSlots ?? defaults.listView.customTableCssSlots,
      customTableCss: params.listView?.customTableCss ?? defaults.listView.customTableCss,
      tableRowStyleRules: params.listView?.tableRowStyleRules ?? defaults.listView.tableRowStyleRules,
      listRowActions:
        params.listView?.listRowActions !== undefined
          ? params.listView.listRowActions
          : defaults.listView.listRowActions,
      ...(params.listView?.viewModePicker === 'tabs' ? { viewModePicker: 'tabs' as const } : {}),
      ...(params.listView?.viewModeDefaultRules?.length
        ? { viewModeDefaultRules: params.listView.viewModeDefaultRules }
        : {}),
    },
    projectManagement: params.projectManagement ?? defaults.projectManagement,
    ...(params.listPageLayout?.sections?.length ? { listPageLayout: params.listPageLayout } : {}),
    ...(params.formManager ? { formManager: params.formManager } : {}),
  };
}
