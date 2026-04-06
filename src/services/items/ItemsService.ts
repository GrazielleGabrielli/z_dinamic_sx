import { getSP } from '../core/sp';
import type { IFieldMetadata } from '../shared/types';
import {
  IItemsQueryOptions,
  IFieldConfig,
  IFilterConfig,
  ISortConfig,
  IPagedResult,
} from './types';

const EXPANDABLE_TYPES: Array<IFieldMetadata['MappedType']> = ['lookup', 'lookupmulti', 'user', 'usermulti'];

const listRef = (sp: ReturnType<typeof getSP>, titleOrId: string) => {
  const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(titleOrId);
  return isGuid
    ? sp.web.lists.getById(titleOrId)
    : sp.web.lists.getByTitle(titleOrId);
};

function normalizeSelectExpand(
  select: string[] | undefined,
  expand: string[] | undefined,
  fieldMetadata: IFieldMetadata[]
): { select: string[]; expand: string[] } {
  const byName = new Map(fieldMetadata.map((f) => [f.InternalName, f]));
  const newSelect: string[] = [];
  const expandArr: string[] = expand ? expand.slice() : [];

  (select ?? []).forEach((f) => {
    if (f.indexOf('/') !== -1) {
      newSelect.push(f);
      return;
    }
    const meta = byName.get(f);
    if (meta && EXPANDABLE_TYPES.indexOf(meta.MappedType) !== -1) {
      if (expandArr.indexOf(f) === -1) expandArr.push(f);
      newSelect.push(`${f}/Id`, `${f}/Title`);
    } else {
      newSelect.push(f);
    }
  });

  return { select: newSelect, expand: expandArr };
}

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
      const { fieldMetadata, ...rest } = options;
      const opts = fieldMetadata?.length
        ? { ...rest, ...normalizeSelectExpand(options.select, options.expand, fieldMetadata) }
        : rest;

      let query = listRef(this.sp, listTitleOrId).items as any;
      if (opts.select?.length) query = query.select(...opts.select);
      if (opts.expand?.length) query = query.expand(...opts.expand);
      if (opts.filter) query = query.filter(opts.filter);
      if (opts.orderBy) query = query.orderBy(opts.orderBy.field, opts.orderBy.ascending);
      if (opts.top) query = query.top(opts.top);
      if (opts.skip) query = query.skip(opts.skip);

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
      const { fieldMetadata, ...rest } = options;
      const opts = fieldMetadata?.length
        ? { ...rest, ...normalizeSelectExpand(options.select, options.expand, fieldMetadata) }
        : rest;

      const top = opts.top ?? pageSize;
      let query = listRef(this.sp, listTitleOrId).items as any;
      if (opts.select?.length) query = query.select(...opts.select);
      if (opts.expand?.length) query = query.expand(...opts.expand);
      if (opts.filter) query = query.filter(opts.filter);
      if (opts.orderBy) query = query.orderBy(opts.orderBy.field, opts.orderBy.ascending);
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
    options: Pick<IItemsQueryOptions, 'select' | 'expand' | 'fieldMetadata'> = {}
  ): Promise<T> {
    try {
      const { fieldMetadata, select, expand } = options;
      const normalized = fieldMetadata?.length
        ? normalizeSelectExpand(select, expand, fieldMetadata)
        : { select: select ?? [], expand: expand ?? [] };

      let query = listRef(this.sp, listTitleOrId).items.getById(itemId) as any;
      if (normalized.select.length) query = query.select(...normalized.select);
      if (normalized.expand.length) query = query.expand(...normalized.expand);

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

  async updateItem(
    listTitleOrId: string,
    itemId: number,
    values: Record<string, unknown>
  ): Promise<void> {
    try {
      await listRef(this.sp, listTitleOrId).items.getById(itemId).update(values);
    } catch (e) {
      throw new Error(`ItemsService.updateItem("${listTitleOrId}", ${itemId}): ${e}`);
    }
  }

  async addItem(listTitleOrId: string, values: Record<string, unknown>): Promise<number> {
    try {
      const result = await listRef(this.sp, listTitleOrId).items.add(values);
      const id = (result as { data?: { Id?: number } })?.data?.Id;
      if (typeof id !== 'number') {
        throw new Error('Resposta sem Id');
      }
      return id;
    } catch (e) {
      throw new Error(`ItemsService.addItem("${listTitleOrId}"): ${e}`);
    }
  }

  async deleteItem(listTitleOrId: string, itemId: number): Promise<void> {
    try {
      await listRef(this.sp, listTitleOrId).items.getById(itemId).delete();
    } catch (e) {
      throw new Error(`ItemsService.deleteItem("${listTitleOrId}", ${itemId}): ${e}`);
    }
  }

  async countItems(listTitleOrId: string, filter: string): Promise<number> {
    try {
      const items = await listRef(this.sp, listTitleOrId).items.filter(filter).select('Id').top(5000)();
      return Array.isArray(items) ? items.length : 0;
    } catch (e) {
      throw new Error(`ItemsService.countItems("${listTitleOrId}"): ${e}`);
    }
  }
}
