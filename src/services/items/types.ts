import { IPagedResult } from '../shared/types';

export { IPagedResult };

export interface IItemsQueryOptions {
  select?: string[];
  expand?: string[];
  filter?: string;
  orderBy?: { field: string; ascending: boolean };
  top?: number;
  skip?: number;
}

export interface IFieldConfig {
  internalName: string;
  expand?: boolean;
  expandFields?: string[];
}

export interface IFilterConfig {
  field: string;
  operator: FilterOperator;
  value: string | number | boolean;
}

export type FilterOperator = 'eq' | 'ne' | 'lt' | 'le' | 'gt' | 'ge' | 'startswith' | 'substringof';

export interface ISortConfig {
  field: string;
  ascending: boolean;
}
