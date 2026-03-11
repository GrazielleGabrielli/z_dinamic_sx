import { IListViewConfig, IListViewFilterConfig, TFilterOperator } from '../config/types';

const NUMERIC_OPERATORS: TFilterOperator[] = ['gt', 'lt', 'ge', 'le'];
const ME_PLACEHOLDER = '[Me]';

function buildFilterSegment(f: IListViewFilterConfig, replaceMe?: string | number): string {
  if (!f.field.trim() || f.value === undefined || f.value === null) return '';
  let value = f.value;
  if (value === ME_PLACEHOLDER && replaceMe !== undefined) {
    value = String(replaceMe);
  }
  const isNumeric = !isNaN(Number(value));
  const val =
    NUMERIC_OPERATORS.indexOf(f.operator) !== -1 && isNumeric
      ? String(value)
      : `'${String(value).replace(/'/g, "''")}'`;

  if (f.operator === 'contains') {
    return `substringof(${val}, ${f.field})`;
  }
  return `${f.field} ${f.operator} ${val}`;
}

export function buildListFilter(
  filters: IListViewFilterConfig[],
  options?: { replaceMe?: string | number }
): string | undefined {
  const replaceMe = options?.replaceMe;
  const segments = filters
    .map((f) => buildFilterSegment(f, replaceMe))
    .filter((s) => s.length > 0);
  if (segments.length === 0) return undefined;
  return segments.join(' and ');
}

export function getActiveViewModeFilters(listView: IListViewConfig): IListViewFilterConfig[] {
  const modes = listView.viewModes;
  if (!modes || modes.length === 0) return listView.filters ?? [];
  const id = listView.activeViewModeId ?? modes[0].id;
  let mode = modes[0];
  for (let i = 0; i < modes.length; i++) {
    if (modes[i].id === id) { mode = modes[i]; break; }
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
}

export function buildListQuery(
  listView: IListViewConfig,
  options?: IBuildListQueryOptions
): IListQueryOptions {
  const select = buildListSelect(listView.columns);
  const expand = buildListExpand(listView.columns);
  const viewModeFilters = getActiveViewModeFilters(listView);
  const filter = buildListFilter(viewModeFilters, { replaceMe: options?.replaceMe });
  const orderBy =
    listView.sort && listView.sort.field
      ? { field: listView.sort.field, ascending: listView.sort.ascending }
      : undefined;

  return { select, expand: expand.length ? expand : undefined, filter, orderBy };
}
