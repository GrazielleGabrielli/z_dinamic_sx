import { TSourceKind, TViewMode, IDynamicViewConfig, TDashboardType, TChartType } from '../../core/config/types';

export interface IWizardFormState {
  kind: TSourceKind;
  title: string;
  mode: TViewMode;
  dashboardEnabled: boolean;
  dashboardType: TDashboardType;
  cardsCount: number;
  chartType: TChartType;
  paginationEnabled: boolean;
  pageSize: number;
  pageSizeOptions: number[];
}

export const WIZARD_INITIAL_STATE: IWizardFormState = {
  kind: 'list',
  title: '',
  mode: 'list',
  dashboardEnabled: false,
  dashboardType: 'cards',
  cardsCount: 3,
  chartType: 'bar',
  paginationEnabled: true,
  pageSize: 20,
  pageSizeOptions: [5, 10, 20, 50, 100],
};

export const PAGE_SIZE_OPTIONS = [5, 10, 20, 50, 100];

export function configToWizardState(config: IDynamicViewConfig): IWizardFormState {
  const cardsCount =
    config.dashboard.cards.length > 0
      ? config.dashboard.cards.length
      : config.dashboard.cardsCount || 3;
  return {
    kind: config.dataSource.kind,
    title: config.dataSource.title,
    mode: config.mode,
    dashboardEnabled: config.dashboard.enabled,
    dashboardType: config.dashboard.dashboardType ?? 'cards',
    cardsCount,
    chartType: config.dashboard.chartType ?? 'bar',
    paginationEnabled: config.pagination.enabled,
    pageSize: config.pagination.pageSize,
    pageSizeOptions: config.pagination.pageSizeOptions,
  };
}
