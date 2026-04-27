import type { IFieldMetadata } from '../../../../services/shared/types';
import { IListViewConfig, IListViewFilterConfig, TFilterOperator } from '../config/types';
import type { IDynamicContext } from '../dynamicTokens/types';
import { resolveObjectTokens } from '../dynamicTokens';
import { isNoteFieldPath } from './fieldQueryRestrictions';

const NUMERIC_OPERATORS: TFilterOperator[] = ['gt', 'lt', 'ge', 'le'];
const ODATA_OPERATORS: TFilterOperator[] = ['eq', 'ne', 'gt', 'lt', 'ge', 'le', 'contains'];

function normalizeOperator(op: unknown): TFilterOperator {
  const s = String(op).toLowerCase();
  for (let i = 0; i < ODATA_OPERATORS.length; i++) {
    if (ODATA_OPERATORS[i] === s) return ODATA_OPERATORS[i];
  }
  return 'eq';
}

function looksLikeToken(value: unknown): boolean {
  if (typeof value !== 'string') return false;
  const s = value.trim();
  return s.length >= 3 && s.charAt(0) === '[' && s.charAt(s.length - 1) === ']';
}

function buildFilterSegment(f: IListViewFilterConfig): string {
  if (!f.field.trim()) return '';
  const value = f.value;
  if (value === undefined) return '';
  if (typeof value === 'string' && looksLikeToken(value)) return '';
  const op = normalizeOperator(f.operator);
  if (value === null) {
    return `${f.field} ${op} null`;
  }
  const isNumeric = !isNaN(Number(value));
  const val =
    NUMERIC_OPERATORS.indexOf(op) !== -1 && isNumeric
      ? String(value)
      : `'${String(value).replace(/'/g, "''")}'`;

  if (op === 'contains') {
    return `substringof(${val}, ${f.field})`;
  }
  return `${f.field} ${op} ${val}`;
}

export interface IBuildListFilterOptions {
  replaceMe?: string | number;
  dynamicContext?: IDynamicContext;
  fieldsMetadata?: IFieldMetadata[];
}

export function buildListFilter(
  filters: IListViewFilterConfig[],
  options?: IBuildListFilterOptions
): string | undefined {
  let resolved = filters;
  if (options?.dynamicContext) {
    try {
      resolved = resolveObjectTokens(filters.slice(), options.dynamicContext) as IListViewFilterConfig[];
    } catch (_) {
      resolved = filters;
    }
  } else if (options?.replaceMe !== undefined) {
    resolved = filters.map((f) => {
      if (f.value === '[Me]' || (typeof f.value === 'string' && f.value.trim().toLowerCase() === '[me]')) {
        return { ...f, value: String(options.replaceMe) };
      }
      return f;
    });
  }
  const segments = resolved
    .map((f) =>
      options?.fieldsMetadata?.length && isNoteFieldPath(f.field, options.fieldsMetadata) ? '' : buildFilterSegment(f)
    )
    .filter((s) => s.length > 0);
  if (segments.length === 0) return undefined;
  return segments.join(' and ');
}

export function getActiveViewModeFilters(listView: IListViewConfig): IListViewFilterConfig[] {
  const modes = listView.viewModes;
  if (!modes || modes.length === 0) return listView.filters ?? [];
  const id = listView.activeViewModeId ?? modes[0].id;
  return getViewModeFiltersById(listView, id);
}

export function getViewModeFiltersById(
  listView: IListViewConfig | undefined,
  modeId: string | undefined
): IListViewFilterConfig[] {
  if (!listView) return [];
  const modes = listView.viewModes;
  if (!modes || modes.length === 0) return listView.filters ?? [];
  const id = modeId ?? listView.activeViewModeId ?? modes[0].id;
  let mode = modes[0];
  for (let i = 0; i < modes.length; i++) {
    if (modes[i].id === id) {
      mode = modes[i];
      break;
    }
  }
  return mode.filters ?? [];
}

export function buildListSelect(columns: IListViewConfig['columns']): string[] {
  if (!columns || columns.length === 0) return ['Id', 'Title'];
  const select: string[] = ['Id'];
  for (const c of columns) {
    if (!c.field) continue;
    if (c.expandField) {
      select.push(`${c.field}/Id`, `${c.field}/Title`);
    } else {
      select.push(c.field);
    }
  }
  if (select.length === 1) select.push('Title');
  return select;
}

function buildListExpand(columns: IListViewConfig['columns']): string[] {
  if (!columns) return [];
  return columns.filter((c) => c.expandField).map((c) => c.field);
}

export interface IListQueryOptions {
  select: string[];
  expand?: string[];
  filter?: string;
  orderBy?: { field: string; ascending: boolean };
}

export interface IBuildListQueryOptions {
  replaceMe?: string | number;
  dynamicContext?: IDynamicContext;
  fieldsMetadata?: IFieldMetadata[];
}

export function buildListQuery(
  listView: IListViewConfig,
  options?: IBuildListQueryOptions
): IListQueryOptions {
  const select = buildListSelect(listView.columns);
  const expand = buildListExpand(listView.columns);
  const viewModeFilters = getActiveViewModeFilters(listView);
  const filter = buildListFilter(viewModeFilters, {
    replaceMe: options?.replaceMe,
    dynamicContext: options?.dynamicContext,
    fieldsMetadata: options?.fieldsMetadata,
  });
  const sortField = listView.sort?.field;
  const orderByBlocked =
    !sortField ||
    (options?.fieldsMetadata?.length && isNoteFieldPath(sortField, options.fieldsMetadata));
  const orderBy =
    listView.sort && sortField && !orderByBlocked
      ? { field: sortField, ascending: listView.sort.ascending }
      : undefined;

  return { select, expand: expand.length ? expand : undefined, filter, orderBy };
}
