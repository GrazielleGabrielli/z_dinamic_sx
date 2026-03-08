import { ItemsService } from '../../../../services';
import { IDashboardCardConfig, IDashboardConfig, IDataSourceConfig, TFilterOperator } from '../config/types';
import { generateDefaultCards } from '../config/utils';
import { IDashboardCardResult, TCardStatus } from './types';

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
