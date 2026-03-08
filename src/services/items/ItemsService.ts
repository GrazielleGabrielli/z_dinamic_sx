import { getSP } from '../core/sp';
import {
  IItemsQueryOptions,
  IFieldConfig,
  IFilterConfig,
  ISortConfig,
  IPagedResult,
} from './types';

const listRef = (sp: ReturnType<typeof getSP>, titleOrId: string) => {
  const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(titleOrId);
  return isGuid
    ? sp.web.lists.getById(titleOrId)
    : sp.web.lists.getByTitle(titleOrId);
};

export class ItemsService {
  private get sp() { return getSP(); }

  /** Monta select e expand a partir de IFieldConfig[] */
  buildSelectExpand(fieldsConfig: IFieldConfig[]): { select: string[]; expand: string[] } {
    const select: string[] = ['Id'];
    const expand: string[] = [];

    for (const fc of fieldsConfig) {
      if (fc.expand && fc.expandFields?.length) {
        expand.push(fc.internalName);
        fc.expandFields.forEach(ef => select.push(`${fc.internalName}/${ef}`));
      } else {
        select.push(fc.internalName);
      }
    }

    return { select, expand };
  }

  /** Constrói string de filtro OData a partir de IFilterConfig[] */
  buildFilter(filters: IFilterConfig[]): string {
    return filters
      .map(f => {
        const val = typeof f.value === 'string' ? `'${f.value}'` : String(f.value);
        if (f.operator === 'startswith') return `startswith(${f.field}, ${val})`;
        if (f.operator === 'substringof') return `substringof(${val}, ${f.field})`;
        return `${f.field} ${f.operator} ${val}`;
      })
      .join(' and ');
  }

  async getItems<T = Record<string, unknown>>(
    listTitleOrId: string,
    options: IItemsQueryOptions = {}
  ): Promise<T[]> {
    try {
      let query = listRef(this.sp, listTitleOrId).items as any;

      if (options.select?.length) query = query.select(...options.select);
      if (options.expand?.length) query = query.expand(...options.expand);
      if (options.filter) query = query.filter(options.filter);
      if (options.orderBy) query = query.orderBy(options.orderBy.field, options.orderBy.ascending);
      if (options.top) query = query.top(options.top);
      if (options.skip) query = query.skip(options.skip);

      return await query() as T[];
    } catch (e) {
      throw new Error(`ItemsService.getItems("${listTitleOrId}"): ${e}`);
    }
  }

  async getPagedItems<T = Record<string, unknown>>(
    listTitleOrId: string,
    options: IItemsQueryOptions = {},
    pageSize = 30,
    skip = 0
  ): Promise<IPagedResult<T>> {
    try {
      const top = options.top ?? pageSize;

      let query = listRef(this.sp, listTitleOrId).items as any;

      if (options.select?.length) query = query.select(...options.select);
      if (options.expand?.length) query = query.expand(...options.expand);
      if (options.filter) query = query.filter(options.filter);
      if (options.orderBy) query = query.orderBy(options.orderBy.field, options.orderBy.ascending);

      query = query.top(top + 1).skip(skip);

      const items: T[] = await query();
      const hasNext = items.length > top;

      return {
        items: hasNext ? items.slice(0, top) : items,
        hasNext,
        nextSkip: skip + top,
      };
    } catch (e) {
      throw new Error(`ItemsService.getPagedItems("${listTitleOrId}"): ${e}`);
    }
  }

  async getItemById<T = Record<string, unknown>>(
    listTitleOrId: string,
    itemId: number,
    options: Pick<IItemsQueryOptions, 'select' | 'expand'> = {}
  ): Promise<T> {
    try {
      let query = listRef(this.sp, listTitleOrId).items.getById(itemId) as any;

      if (options.select?.length) query = query.select(...options.select);
      if (options.expand?.length) query = query.expand(...options.expand);

      return await query() as T;
    } catch (e) {
      throw new Error(`ItemsService.getItemById("${listTitleOrId}", ${itemId}): ${e}`);
    }
  }

  /** Aplica múltiplos filtros ao options existente */
  applyFilter(options: IItemsQueryOptions, filters: IFilterConfig[]): IItemsQueryOptions {
    const filterStr = this.buildFilter(filters);
    return {
      ...options,
      filter: options.filter ? `(${options.filter}) and (${filterStr})` : filterStr,
    };
  }

  applyOrderBy(options: IItemsQueryOptions, sort: ISortConfig): IItemsQueryOptions {
    return { ...options, orderBy: sort };
  }
}
