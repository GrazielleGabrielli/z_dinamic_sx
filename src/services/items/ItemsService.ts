import { fileFromServerRelativePath } from '@pnp/sp/files';

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

function coerceListItemId(v: unknown): number | undefined {
  if (typeof v === 'number' && isFinite(v)) return v;
  if (typeof v === 'string' && v.trim() !== '') {
    const n = parseInt(v, 10);
    return isNaN(n) ? undefined : n;
  }
  return undefined;
}

/** Resposta de `items.add` no PnP v4: corpo OData com `Id` na raiz; formatos antigos usam `d` ou `data`. */
function extractCreatedItemId(result: unknown): number | undefined {
  if (result === null || typeof result !== 'object') return undefined;
  const r = result as Record<string, unknown>;
  const top = coerceListItemId(r.Id ?? r.id);
  if (top !== undefined) return top;
  const data = r.data;
  if (data && typeof data === 'object') {
    const id = coerceListItemId((data as Record<string, unknown>).Id ?? (data as Record<string, unknown>).id);
    if (id !== undefined) return id;
  }
  const d = r.d;
  if (d && typeof d === 'object') {
    const id = coerceListItemId((d as Record<string, unknown>).Id ?? (d as Record<string, unknown>).id);
    if (id !== undefined) return id;
  }
  return undefined;
}

const FILE_LIBRARY_BASE_TEMPLATES = new Set([101, 109]);

const READ_ONLY_ITEM_KEYS_FOR_FILE_UPLOAD = new Set([
  'Id',
  'ID',
  'GUID',
  'UniqueId',
  'FileLeafRef',
  'FileRef',
  'FileDirRef',
  'File_x0020_Type',
  'CheckoutUser',
  'CheckedOutUserId',
  'SyncClientId',
  'ServerRedirectedEmbedUrl',
  'ServerRedirectedEmbedUri',
  'ParentUniqueId',
  'ScopeId',
]);

const COMPUTED_LINK_FIELDS = new Set([
  'LinkFilename',
  'LinkFilenameNoMenu',
  'LinkTitle',
  'LinkTitleNoMenu',
]);

function sanitizePayloadForFileLibraryItem(values: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  Object.keys(values).forEach((k) => {
    if (READ_ONLY_ITEM_KEYS_FOR_FILE_UPLOAD.has(k)) return;
    if (COMPUTED_LINK_FIELDS.has(k)) return;
    if (k.indexOf('OData__') === 0) return;
    out[k] = values[k];
  });
  return out;
}

function defaultPlaceholderFileName(): string {
  return `registro-${Date.now()}-${Math.random().toString(36).slice(2, 9)}.txt`;
}

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

  async addItem(
    listTitleOrId: string,
    values: Record<string, unknown>,
    primaryFiles?: File[]
  ): Promise<{ id: number; filesForAttachments: File[] }> {
    try {
      const list = listRef(this.sp, listTitleOrId);
      const listInfo = await list.select('BaseTemplate')();
      const baseTemplate = (listInfo as { BaseTemplate?: number }).BaseTemplate;
      const uploaded = primaryFiles ?? [];

      if (baseTemplate !== undefined && FILE_LIBRARY_BASE_TEMPLATES.has(baseTemplate)) {
        const first = uploaded.length ? uploaded[0] : undefined;
        const fileName = first?.name?.trim() ? first.name.trim() : defaultPlaceholderFileName();
        const body: Blob | string = first ?? '\uFEFF';

        const fileInfo = await list.rootFolder.files.addUsingPath(fileName, body, {
          EnsureUniqueFileName: true,
        });
        const rel = (fileInfo as { ServerRelativeUrl?: string }).ServerRelativeUrl;
        if (!rel) {
          throw new Error('Upload sem ServerRelativeUrl');
        }

        const file = fileFromServerRelativePath(this.sp.web, rel);
        const item = await file.getItem<{ Id?: number }>('Id');
        const id = coerceListItemId(item.Id);
        if (id === undefined) {
          throw new Error('Resposta sem Id');
        }

        const meta = sanitizePayloadForFileLibraryItem(values);
        if (Object.keys(meta).length) {
          await list.items.getById(id).update(meta);
        }

        const filesForAttachments = first ? uploaded.slice(1) : uploaded;
        return { id, filesForAttachments };
      }

      const result = await list.items.add(values);
      const newId = extractCreatedItemId(result);
      if (newId === undefined) {
        throw new Error('Resposta sem Id');
      }
      return { id: newId, filesForAttachments: uploaded };
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

  async countItems(listTitleOrId: string, filter?: string): Promise<number> {
    try {
      const f = (filter ?? '').trim();
      let query = listRef(this.sp, listTitleOrId).items.select('Id').top(5000);
      if (f) {
        query = query.filter(f);
      }
      const items = await query();
      return Array.isArray(items) ? items.length : 0;
    } catch (e) {
      throw new Error(`ItemsService.countItems("${listTitleOrId}"): ${e}`);
    }
  }

  async getItemVersions(
    listTitleOrId: string,
    itemId: number
  ): Promise<
    { versionLabel: string; versionId: number; created?: string; isCurrentVersion?: boolean }[]
  > {
    try {
      const raw = await listRef(this.sp, listTitleOrId)
        .items.getById(itemId)
        .versions.select('VersionLabel', 'VersionId', 'Created', 'IsCurrentVersion')();
      const rows = Array.isArray(raw) ? raw : [];
      const out: { versionLabel: string; versionId: number; created?: string; isCurrentVersion?: boolean }[] = [];
      for (let i = 0; i < rows.length; i++) {
        const r = rows[i] as Record<string, unknown>;
        const versionId = coerceListItemId(r.VersionId ?? r.versionId);
        const vl = r.VersionLabel ?? r.versionLabel;
        const versionLabel = typeof vl === 'string' && vl.trim() ? vl.trim() : String(versionId ?? i + 1);
        if (versionId === undefined) continue;
        const cr = r.Created ?? r.created;
        const created = typeof cr === 'string' ? cr : cr instanceof Date ? cr.toISOString() : undefined;
        const icv = r.IsCurrentVersion ?? r.isCurrentVersion;
        const isCurrentVersion = icv === true || icv === 1;
        out.push({ versionLabel, versionId, created, isCurrentVersion });
      }
      out.sort((a, b) => b.versionId - a.versionId);
      return out;
    } catch (e) {
      throw new Error(`ItemsService.getItemVersions("${listTitleOrId}", ${itemId}): ${e}`);
    }
  }
}
