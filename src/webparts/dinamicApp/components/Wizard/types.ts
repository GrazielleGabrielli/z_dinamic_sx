import { TSourceKind, TViewMode, IDynamicViewConfig } from '../../core/config/types';

export interface IWizardFormState {
  kind: TSourceKind;
  title: string;
  mode: TViewMode;
  dashboardEnabled: boolean;
  cardsCount: number;
  paginationEnabled: boolean;
  pageSize: number;
  pageSizeOptions: number[];
}

export const WIZARD_INITIAL_STATE: IWizardFormState = {
  kind: 'list',
  title: '',
  mode: 'list',
  dashboardEnabled: false,
  cardsCount: 3,
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
    cardsCount,
    paginationEnabled: config.pagination.enabled,
    pageSize: config.pagination.pageSize,
    pageSizeOptions: config.pagination.pageSizeOptions,
  };
}
