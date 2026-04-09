import {
  IDynamicViewConfig,
  IDashboardCardConfig,
  TAggregateType,
  IFormManagerConfig,
  TViewMode,
} from '../types';
import { FORM_OCULTOS_STEP_ID } from '../types/formManager';
import { getDefaultDashboardCardStyle } from '../../dashboard/utils';

export function getDefaultFormManagerConfig(): IFormManagerConfig {
  return {
    sections: [
      { id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', visible: true },
      { id: 'main', title: 'Geral', visible: true },
    ],
    fields: [],
    rules: [],
    steps: [
      { id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', fieldNames: [] },
      { id: 'main', title: 'Geral', fieldNames: [] },
    ],
    stepLayout: 'segmented',
  };
}

export function getDefaultConfigForMode(mode: TViewMode): IDynamicViewConfig {
  const base = getDefaultConfig();
  return {
    ...base,
    mode,
    ...(mode === 'formManager' ? { formManager: getDefaultFormManagerConfig() } : {}),
  };
}

export function getDefaultConfig(): IDynamicViewConfig {
  return {
    dataSource: {
      kind: 'list',
      title: '',
    },
    mode: 'list',
    dashboard: {
      enabled: false,
      dashboardType: 'cards',
      cardsCount: 0,
      cards: [],
      chartType: 'bar',
    },
    pagination: {
      enabled: true,
      pageSize: 10,
      pageSizeOptions: [5, 10, 20, 50, 100],
      layout: 'buttons',
    },
    listView: {
      columns: [{ field: 'Title' }],
      filters: [],
      sort: null,
      viewModes: [
        { id: 'all', label: 'Todas', filters: [] },
        { id: 'mine', label: 'Minhas', filters: [{ field: 'Author/Id', operator: 'eq', value: '[Me]' }] },
      ],
      activeViewModeId: 'all',
      listCardViewEnabled: false,
    },
    projectManagement: {
      columns: [],
    },
  };
}

export function isConfigured(config: IDynamicViewConfig): boolean {
  return config.dataSource.title.trim().length > 0;
}

export function generateDefaultCards(count: number): IDashboardCardConfig[] {
  const defaultStyle = getDefaultDashboardCardStyle();
  const cards: IDashboardCardConfig[] = [];
  for (let i = 0; i < count; i++) {
    cards.push({
      id: `card_${i + 1}`,
      title: `Card ${i + 1}`,
      subtitle: '',
      aggregate: 'count' as TAggregateType,
      emptyValueText: 'Nenhum item',
      errorText: 'Erro ao carregar',
      loadingText: 'Carregando...',
      style: { ...defaultStyle },
    });
  }
  return cards;
}
