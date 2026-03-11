import { IDynamicViewConfig, IDashboardCardConfig, TAggregateType } from '../types';
import { getDefaultDashboardCardStyle } from '../../dashboard/utils';

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
      columns: [],
      filters: [],
      sort: null,
      viewModes: [
        { id: 'all', label: 'Todas', filters: [] },
        { id: 'mine', label: 'Minhas', filters: [{ field: 'Author/Id', operator: 'eq', value: '[Me]' }] },
      ],
      activeViewModeId: 'all',
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
