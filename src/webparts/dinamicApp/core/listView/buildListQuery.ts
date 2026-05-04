import type { IFieldMetadata, FieldMappedType } from '../../../../services/shared/types';
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

/**
 * Tipos de campo que devem usar substringof/contains em vez de eq.
 * Para choice/boolean/number usamos eq direto; para text/user/lookup usamos substringof.
 */
const SUBSTRING_TYPES: FieldMappedType[] = ['text', 'multiline', 'url', 'user', 'usermulti', 'lookup', 'lookupmulti'];

/**
 * Constrói OData para os filtros da barra de filtros da tabela (tableFilterFields).
 * @param values - mapa { fieldName: valorDigitado }
 * @param fieldMeta - metadados dos campos para determinar o tipo
 */
export function buildTableTopFiltersOData(
  values: Record<string, string>,
  fieldMeta: IFieldMetadata[]
): string | undefined {
  const metaByName = new Map(fieldMeta.map((m) => [m.InternalName, m]));
  const parts: string[] = [];
  for (const field in values) {
    if (!Object.prototype.hasOwnProperty.call(values, field)) continue;
    const raw = (values[field] ?? '').trim();
    if (!raw) continue;

    const isExpandPath = field.indexOf('/') !== -1;
    const baseName = isExpandPath ? field.split('/')[0] : field;
    const meta = metaByName.get(baseName);
    const mappedType: FieldMappedType = meta?.MappedType ?? 'text';

    if (mappedType === 'boolean') {
      const lower = raw.toLowerCase();
      const boolVal = lower === 'true' || lower === '1' || lower === 'sim' ? 'true' : 'false';
      parts.push(`${field} eq ${boolVal}`);
    } else if (mappedType === 'number' || mappedType === 'currency') {
      const n = Number(raw);
      if (!isNaN(n)) parts.push(`${field} eq ${n}`);
    } else if (mappedType === 'choice' || mappedType === 'multichoice') {
      parts.push(`${field} eq '${raw.replace(/'/g, "''")}'`);
    } else if (mappedType === 'datetime') {
      const d = new Date(raw);
      if (!isNaN(d.getTime())) {
        parts.push(`${field} ge datetime'${d.toISOString()}'`);
      }
    } else if (SUBSTRING_TYPES.indexOf(mappedType) !== -1) {
      parts.push(`substringof('${raw.replace(/'/g, "''")}', ${field})`);
    } else {
      parts.push(`substringof('${raw.replace(/'/g, "''")}', ${field})`);
    }
  }
  return parts.length > 0 ? parts.join(' and ') : undefined;
}

export function buildListSelect(columns: IListViewConfig['columns']): string[] {
  if (!columns || columns.length === 0) return ['Id', 'Title'];
  const selectSet = new Set<string>(['Id']);
  for (const c of columns) {
    if (!c.field) continue;
    if (c.expandField) {
      const df = c.expandField.trim() || 'Title';
      const root = c.field;
      if (df.indexOf('/') !== -1) {
        const idSuffix = [...df.split('/').slice(0, -1), 'Id'].join('/');
        selectSet.add(`${root}/${idSuffix}`);
        selectSet.add(`${root}/${df}`);
      } else {
        selectSet.add(`${root}/Id`);
        selectSet.add(`${root}/${df}`);
      }
    } else {
      selectSet.add(c.field);
    }
  }
  const select = Array.from(selectSet);
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
