import { ItemsService } from '../../../../services';
import { IDashboardCardConfig, IDashboardConfig, IDataSourceConfig, IChartSeriesConfig, TFilterOperator } from '../config/types';
import { generateDefaultCards } from '../config/utils';
import { IDashboardCardResult, IChartSeriesResult, TCardStatus } from './types';

const NUMERIC_OPERATORS: TFilterOperator[] = ['gt', 'lt', 'ge', 'le'];

export class DashboardEngine {
  private readonly itemsService = new ItemsService();

  private buildFilterString(filter: IDashboardCardConfig['filter']): string | undefined {
    if (!filter) return undefined;
    const isNumeric = !isNaN(Number(filter.value));
    const val =
      NUMERIC_OPERATORS.indexOf(filter.operator) !== -1 && isNumeric
        ? filter.value
        : `'${filter.value}'`;

    if (filter.operator === 'contains') {
      return `substringof(${val}, ${filter.field})`;
    }
    return `${filter.field} ${filter.operator} ${val}`;
  }

  private cardBase(card: IDashboardCardConfig): Pick<IDashboardCardResult, 'id' | 'title' | 'subtitle' | 'aggregate' | 'style' | 'emptyValueText' | 'errorText' | 'loadingText'> {
    return {
      id: card.id,
      title: card.title,
      subtitle: card.subtitle,
      aggregate: card.aggregate,
      style: card.style,
      emptyValueText: card.emptyValueText,
      errorText: card.errorText,
      loadingText: card.loadingText,
    };
  }

  async computeCard(
    card: IDashboardCardConfig,
    dataSource: IDataSourceConfig
  ): Promise<IDashboardCardResult> {
    const base = this.cardBase(card);

    try {
      const filterStr = this.buildFilterString(card.filter);

      if (card.aggregate === 'count') {
        const items = await this.itemsService.getItems(dataSource.title, {
          select: ['Id'],
          filter: filterStr,
          top: 5000,
        });
        return { ...base, value: items.length, status: 'ready' };
      }

      if (card.aggregate === 'sum') {
        if (!card.field) {
          return { ...base, value: undefined, status: 'error', error: 'Campo não definido para soma' };
        }
        const items = await this.itemsService.getItems<Record<string, unknown>>(
          dataSource.title,
          { select: ['Id', card.field], filter: filterStr, top: 5000 }
        );
        const fieldName = card.field;
        const sum = items.reduce((acc, item) => {
          const raw = Number(item[fieldName]);
          return acc + (isNaN(raw) ? 0 : raw);
        }, 0);
        return { ...base, value: sum, status: 'ready' };
      }

      return { ...base, value: undefined, status: 'error', error: 'Tipo de agregação não suportado' };
    } catch (err) {
      return { ...base, value: undefined, status: 'error', error: String(err) };
    }
  }

  async computeAll(
    config: IDashboardConfig,
    dataSource: IDataSourceConfig
  ): Promise<IDashboardCardResult[]> {
    const cards =
      config.cards.length > 0 ? config.cards : generateDefaultCards(config.cardsCount);

    return Promise.all(cards.map((card) => this.computeCard(card, dataSource)));
  }

  async computeSeries(
    series: IChartSeriesConfig,
    dataSource: IDataSourceConfig
  ): Promise<IChartSeriesResult> {
    try {
      const filterStr = this.buildFilterString(series.filter);

      if (series.aggregate === 'count') {
        const items = await this.itemsService.getItems(dataSource.title, {
          select: ['Id'],
          filter: filterStr,
          top: 5000,
        });
        return { id: series.id, label: series.label, value: items.length, color: series.color, status: 'ready' };
      }

      if (series.aggregate === 'sum') {
        if (!series.field) {
          return { id: series.id, label: series.label, value: 0, color: series.color, status: 'error', error: 'Campo não definido' };
        }
        const items = await this.itemsService.getItems<Record<string, unknown>>(
          dataSource.title,
          { select: ['Id', series.field], filter: filterStr, top: 5000 }
        );
        const fieldName = series.field;
        const sum = items.reduce((acc, item) => {
          const raw = Number(item[fieldName]);
          return acc + (isNaN(raw) ? 0 : raw);
        }, 0);
        return { id: series.id, label: series.label, value: sum, color: series.color, status: 'ready' };
      }

      return { id: series.id, label: series.label, value: 0, color: series.color, status: 'error', error: 'Agregação não suportada' };
    } catch (err) {
      return { id: series.id, label: series.label, value: 0, color: series.color, status: 'error', error: String(err) };
    }
  }

  async computeAllSeries(
    config: IDashboardConfig,
    dataSource: IDataSourceConfig
  ): Promise<IChartSeriesResult[]> {
    const series = config.chartSeries ?? [];
    return Promise.all(series.map((s) => this.computeSeries(s, dataSource)));
  }

  buildLoadingResults(config: IDashboardConfig): IDashboardCardResult[] {
    const cards =
      config.cards.length > 0 ? config.cards : generateDefaultCards(config.cardsCount);

    return cards.map((card) => ({
      ...this.cardBase(card),
      value: undefined,
      status: 'loading' as TCardStatus,
    }));
  }
}
