import { ItemsService } from '../../../../services';
import type { IFieldMetadata } from '../../../../services';
import { IDashboardCardConfig, IDashboardCardFilter, IDashboardConfig, IDataSourceConfig, IChartSeriesConfig, TFilterOperator } from '../config/types';
import { generateDefaultCards } from '../config/utils';
import type { IDynamicContext } from '../dynamicTokens/types';
import { resolveObjectTokens, isDynamicToken } from '../dynamicTokens';
import { IDashboardCardResult, IChartSeriesResult, TCardStatus } from './types';

const NUMERIC_OPERATORS: TFilterOperator[] = ['gt', 'lt', 'ge', 'le'];
const ODATA_OPERATORS: TFilterOperator[] = ['eq', 'ne', 'gt', 'lt', 'ge', 'le', 'contains'];

function normalizeDashboardOperator(op: unknown): TFilterOperator {
  const s = String(op).toLowerCase();
  for (let i = 0; i < ODATA_OPERATORS.length; i++) {
    if (ODATA_OPERATORS[i] === s) return ODATA_OPERATORS[i];
  }
  return 'eq';
}

export class DashboardEngine {
  private readonly itemsService = new ItemsService();

  private lookupMetaForFilter(
    rawField: string,
    fieldMetadata?: IFieldMetadata[]
  ): IFieldMetadata | undefined {
    const f = rawField.trim();
    if (!f || !fieldMetadata?.length) return undefined;
    const byName = new Map(fieldMetadata.map((m) => [m.InternalName, m]));
    const direct = byName.get(f);
    if (direct && (direct.MappedType === 'lookup' || direct.MappedType === 'user')) return direct;
    if (f.length > 2 && f.endsWith('Id')) {
      const base = f.slice(0, -2);
      const meta = byName.get(base);
      if (meta && (meta.MappedType === 'lookup' || meta.MappedType === 'user')) return meta;
    }
    return undefined;
  }

  private buildOneFilterSegment(
    filter: IDashboardCardFilter,
    dynamicContext?: IDynamicContext,
    fieldMetadata?: IFieldMetadata[]
  ): string | undefined {
    if (!filter || !filter.field.trim()) return undefined;
    let resolved: IDashboardCardFilter = filter;
    if (dynamicContext) {
      try {
        const resolvedObj = resolveObjectTokens({ ...filter }, dynamicContext) as IDashboardCardFilter;
        if (!resolvedObj || resolvedObj.value === undefined || (typeof resolvedObj.value === 'string' && isDynamicToken(resolvedObj.value))) {
          return undefined;
        }
        resolved = resolvedObj;
      } catch (_) {
        resolved = filter;
      }
    }
    const op = normalizeDashboardOperator(resolved.operator);
    const value = resolved.value;
    const fieldName = resolved.field.trim();

    const lk = this.lookupMetaForFilter(fieldName, fieldMetadata);
    if (lk) {
      if (op === 'contains') return undefined;
      const idNum = parseInt(String(value).trim(), 10);
      if (Number.isNaN(idNum)) return undefined;
      const idField = `${lk.InternalName}Id`;
      return `${idField} ${op} ${idNum}`;
    }

    const isNumeric = !isNaN(Number(value));
    const val =
      NUMERIC_OPERATORS.indexOf(op) !== -1 && isNumeric
        ? value
        : `'${String(value).replace(/'/g, "''")}'`;

    if (op === 'contains') {
      return `substringof(${val}, ${fieldName})`;
    }
    return `${fieldName} ${op} ${val}`;
  }

  private buildFilterString(
    filterOrFilters: IDashboardCardFilter | IDashboardCardFilter[] | undefined,
    dynamicContext?: IDynamicContext,
    fieldMetadata?: IFieldMetadata[]
  ): string | undefined {
    const list: IDashboardCardFilter[] = !filterOrFilters
      ? []
      : Array.isArray(filterOrFilters)
        ? filterOrFilters
        : [filterOrFilters];
    const segments: string[] = [];
    for (let i = 0; i < list.length; i++) {
      const seg = this.buildOneFilterSegment(list[i], dynamicContext, fieldMetadata);
      if (seg) segments.push(seg);
    }
    if (segments.length === 0) return undefined;
    return segments.join(' and ');
  }

  private combineODataParts(a?: string, b?: string): string | undefined {
    const p = [(a ?? '').trim(), (b ?? '').trim()].filter((s) => s.length > 0);
    return p.length ? p.join(' and ') : undefined;
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
    dataSource: IDataSourceConfig,
    fieldMetadata?: IFieldMetadata[],
    dynamicContext?: IDynamicContext,
    linkedViewModeOData?: string
  ): Promise<IDashboardCardResult> {
    const base = this.cardBase(card);

    try {
      const effectiveFilters = card.filters && card.filters.length > 0 ? card.filters : (card.filter ? [card.filter] : []);
      const built = this.buildFilterString(effectiveFilters, dynamicContext, fieldMetadata);
      const filterStr = this.combineODataParts(built, linkedViewModeOData);
      const baseOptions = {
        filter: filterStr,
        top: 5000,
        fieldMetadata,
        ...(dataSource.webServerRelativeUrl?.trim()
          ? { webServerRelativeUrl: dataSource.webServerRelativeUrl.trim() }
          : {}),
      };

      if (card.aggregate === 'count') {
        const items = await this.itemsService.getItems(dataSource.title, {
          ...baseOptions,
          select: ['Id'],
        });
        return { ...base, value: items.length, status: 'ready' };
      }

      if (card.aggregate === 'sum') {
        if (!card.field) {
          return { ...base, value: undefined, status: 'error', error: 'Campo não definido para soma' };
        }
        const select = card.expandField ? ['Id', `${card.field}/${card.expandField}`] : ['Id', card.field];
        const expand = card.expandField ? [card.field] : undefined;
        const items = await this.itemsService.getItems<Record<string, unknown>>(
          dataSource.title,
          { ...baseOptions, select, expand }
        );
        const fieldName = card.field;
        const expandField = card.expandField;
        const sum = items.reduce((acc, item) => {
          const rawVal = expandField && item[fieldName] && typeof item[fieldName] === 'object'
            ? (item[fieldName] as Record<string, unknown>)[expandField]
            : item[fieldName];
          const raw = Number(rawVal);
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
    dataSource: IDataSourceConfig,
    fieldMetadata?: IFieldMetadata[],
    dynamicContext?: IDynamicContext,
    linkedViewModeOData?: string
  ): Promise<IDashboardCardResult[]> {
    const cards =
      config.cards.length > 0 ? config.cards : generateDefaultCards(config.cardsCount);

    return Promise.all(
      cards.map((card) => this.computeCard(card, dataSource, fieldMetadata, dynamicContext, linkedViewModeOData))
    );
  }

  async computeSeries(
    series: IChartSeriesConfig,
    dataSource: IDataSourceConfig,
    fieldMetadata?: IFieldMetadata[],
    dynamicContext?: IDynamicContext,
    linkedViewModeOData?: string
  ): Promise<IChartSeriesResult> {
    try {
      const effectiveFilters = series.filters && series.filters.length > 0 ? series.filters : (series.filter ? [series.filter] : []);
      const built = this.buildFilterString(effectiveFilters, dynamicContext, fieldMetadata);
      const filterStr = this.combineODataParts(built, linkedViewModeOData);
      const baseOptions = {
        filter: filterStr,
        top: 5000,
        fieldMetadata,
        ...(dataSource.webServerRelativeUrl?.trim()
          ? { webServerRelativeUrl: dataSource.webServerRelativeUrl.trim() }
          : {}),
      };

      if (series.aggregate === 'count') {
        const items = await this.itemsService.getItems(dataSource.title, {
          ...baseOptions,
          select: ['Id'],
        });
        return { id: series.id, label: series.label, value: items.length, color: series.color, status: 'ready' };
      }

      if (series.aggregate === 'sum') {
        if (!series.field) {
          return { id: series.id, label: series.label, value: 0, color: series.color, status: 'error', error: 'Campo não definido' };
        }
        const select = series.expandField ? ['Id', `${series.field}/${series.expandField}`] : ['Id', series.field];
        const expand = series.expandField ? [series.field] : undefined;
        const items = await this.itemsService.getItems<Record<string, unknown>>(
          dataSource.title,
          { ...baseOptions, select, expand }
        );
        const fieldName = series.field;
        const expandField = series.expandField;
        const sum = items.reduce((acc, item) => {
          const rawVal = expandField && item[fieldName] && typeof item[fieldName] === 'object'
            ? (item[fieldName] as Record<string, unknown>)[expandField]
            : item[fieldName];
          const raw = Number(rawVal);
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
    dataSource: IDataSourceConfig,
    fieldMetadata?: IFieldMetadata[],
    dynamicContext?: IDynamicContext,
    linkedViewModeOData?: string
  ): Promise<IChartSeriesResult[]> {
    const series = config.chartSeries ?? [];
    return Promise.all(
      series.map((s) => this.computeSeries(s, dataSource, fieldMetadata, dynamicContext, linkedViewModeOData))
    );
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
