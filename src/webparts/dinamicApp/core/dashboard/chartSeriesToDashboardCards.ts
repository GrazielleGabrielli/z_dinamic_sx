import type { IChartSeriesConfig, IDashboardCardConfig, IDashboardCardFilter } from '../config/types';
import { getDefaultDashboardCardStyle } from './utils/dashboardCardStyles';

const SERIES_COLORS_FROM_CARDS = [
  '#0078d4', '#2b88d8', '#71afe5',
  '#00b294', '#ffaa44', '#d13438',
  '#8764b8', '#038387', '#ca5010',
];

function effectiveCardFilters(c: IDashboardCardConfig): IDashboardCardFilter[] {
  if (c.filters && c.filters.length > 0) {
    return c.filters.map((f) => ({ field: f.field, operator: f.operator, value: f.value }));
  }
  if (c.filter) {
    return [{ field: c.filter.field, operator: c.filter.operator, value: c.filter.value }];
  }
  return [];
}

function cardIdToSeriesId(cardId: string, index: number): string {
  const prefix = 'card_';
  if (cardId.indexOf(prefix) === 0) {
    return 'series_' + cardId.slice(prefix.length);
  }
  const safe = cardId.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 80);
  return safe.length > 0 ? `series_${safe}` : `series_${index + 1}`;
}

export function dashboardCardsToChartSeries(cards: IDashboardCardConfig[]): IChartSeriesConfig[] {
  const out: IChartSeriesConfig[] = [];
  for (let i = 0; i < cards.length; i++) {
    const c = cards[i];
    const filters = effectiveCardFilters(c);
    const s: IChartSeriesConfig = {
      id: cardIdToSeriesId(c.id, i),
      label: (c.title ?? '').trim() || `Série ${i + 1}`,
      aggregate: c.aggregate,
      color: SERIES_COLORS_FROM_CARDS[i % SERIES_COLORS_FROM_CARDS.length],
    };
    if (c.field !== undefined && c.field.trim().length > 0) s.field = c.field.trim();
    if (c.expandField !== undefined && c.expandField.trim().length > 0) s.expandField = c.expandField.trim();
    if (filters.length > 0) s.filters = filters;
    out.push(s);
  }
  return out;
}

function effectiveSeriesFilters(s: IChartSeriesConfig): IDashboardCardFilter[] {
  if (s.filters && s.filters.length > 0) {
    return s.filters.map((f) => ({ field: f.field, operator: f.operator, value: f.value }));
  }
  if (s.filter) {
    return [{ field: s.filter.field, operator: s.filter.operator, value: s.filter.value }];
  }
  return [];
}

function cardTitleFromSeries(s: IChartSeriesConfig, index: number): string {
  const filters = effectiveSeriesFilters(s);
  if (filters.length === 1) {
    const v = String(filters[0].value ?? '').trim();
    if (v.length > 0) return v;
  }
  const lbl = (s.label ?? '').trim();
  if (lbl.length > 0) return lbl;
  return `Card ${index + 1}`;
}

function seriesIdToCardId(seriesId: string, index: number): string {
  const prefix = 'series_';
  if (seriesId.indexOf(prefix) === 0) {
    return 'card_' + seriesId.slice(prefix.length);
  }
  const safe = seriesId.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 80);
  return safe.length > 0 ? `card_${safe}` : `card_${index + 1}`;
}

export function chartSeriesToDashboardCards(series: IChartSeriesConfig[]): IDashboardCardConfig[] {
  const defaultStyle = getDefaultDashboardCardStyle();
  const out: IDashboardCardConfig[] = [];
  for (let i = 0; i < series.length; i++) {
    const s = series[i];
    const filters = effectiveSeriesFilters(s);
    const card: IDashboardCardConfig = {
      id: seriesIdToCardId(s.id, i),
      title: cardTitleFromSeries(s, i),
      subtitle: '',
      aggregate: s.aggregate,
      emptyValueText: 'Nenhum item',
      errorText: 'Erro ao carregar',
      loadingText: 'Carregando...',
      style: { ...defaultStyle },
    };
    if (s.field !== undefined && s.field.trim().length > 0) card.field = s.field.trim();
    if (s.expandField !== undefined && s.expandField.trim().length > 0) card.expandField = s.expandField.trim();
    if (filters.length > 0) card.filters = filters;
    out.push(card);
  }
  return out;
}
