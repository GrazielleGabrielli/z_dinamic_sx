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
    },
    listView: {
      columns: [],
      filters: [],
      sort: null,
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
