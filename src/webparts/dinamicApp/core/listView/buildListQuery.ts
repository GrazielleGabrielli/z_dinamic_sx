import { IListViewConfig, IListViewFilterConfig, TFilterOperator } from '../config/types';

const NUMERIC_OPERATORS: TFilterOperator[] = ['gt', 'lt', 'ge', 'le'];

function buildFilterSegment(f: IListViewFilterConfig): string {
  if (!f.field.trim() || f.value === undefined || f.value === null) return '';
  const isNumeric = !isNaN(Number(f.value));
  const val =
    NUMERIC_OPERATORS.indexOf(f.operator) !== -1 && isNumeric
      ? String(f.value)
      : `'${String(f.value).replace(/'/g, "''")}'`;

  if (f.operator === 'contains') {
    return `substringof(${val}, ${f.field})`;
  }
  return `${f.field} ${f.operator} ${val}`;
}

export function buildListFilter(filters: IListViewFilterConfig[]): string | undefined {
  const segments = filters
    .map(buildFilterSegment)
    .filter((s) => s.length > 0);
  if (segments.length === 0) return undefined;
  return segments.join(' and ');
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

export function buildListQuery(listView: IListViewConfig): IListQueryOptions {
  const select = buildListSelect(listView.columns);
  const expand = buildListExpand(listView.columns);
  const filter = buildListFilter(listView.filters);
  const orderBy =
    listView.sort && listView.sort.field
      ? { field: listView.sort.field, ascending: listView.sort.ascending }
      : undefined;

  return { select, expand: expand.length ? expand : undefined, filter, orderBy };
}
