import type { ReactNode } from 'react';

export type InternalFieldType =
  | 'text'
  | 'note'
  | 'number'
  | 'currency'
  | 'date'
  | 'boolean'
  | 'choice'
  | 'multiChoice'
  | 'lookup'
  | 'lookupMulti'
  | 'user'
  | 'userMulti'
  | 'url'
  | 'file'
  | 'managedMetadata'
  | 'calculated'
  | 'image'
  | 'unknown';

export type TTableColumnAlign = 'left' | 'center' | 'right';

export type TTableColumnRenderMode =
  | 'default'
  | 'badge'
  | 'link'
  | 'avatar'
  | 'date'
  | 'custom';

export type TTableColumnSourceKind =
  | 'currentListField'
  | 'lookupExpand'
  | 'externalListField'
  | 'computed';

export interface ITableColumnExpandConfig {
  displayField?: string;
  displayFields?: string[];
  separator?: string;
  nestedPath?: string;
}

export interface ITableColumnFormattingConfig {
  emptyValueText?: string;
  trueText?: string;
  falseText?: string;
  dateFormat?: string;
  numberDecimals?: number;
  currencyCode?: string;
  prefix?: string;
  suffix?: string;
}

export interface ITableColumnSourceConfig {
  kind: TTableColumnSourceKind;
  listTitle?: string;
  fieldInternalName?: string;
  relationField?: string;
  foreignKeyField?: string;
}

export interface ITableColumnConfig {
  id: string;
  internalName: string;
  label: string;
  visible: boolean;
  sortable: boolean;
  fieldType?: InternalFieldType;
  width?: string;
  minWidth?: number;
  maxWidth?: number;
  align?: TTableColumnAlign;
  fallbackValue?: string;
  renderMode?: TTableColumnRenderMode;
  expandConfig?: ITableColumnExpandConfig;
  formatting?: ITableColumnFormattingConfig;
  source?: ITableColumnSourceConfig;
}

export interface ISortConfig {
  field: string;
  direction: 'asc' | 'desc';
}

export interface ITableConfig {
  enabled: boolean;
  columns: ITableColumnConfig[];
  sortable: boolean;
  defaultSort?: ISortConfig;
  allowColumnToggle?: boolean;
  allowColumnReorder?: boolean;
  stickyHeader?: boolean;
  dense?: boolean;
  emptyMessage?: string;
}

export interface ITableDataRequest {
  select: string[];
  expand?: string[];
  orderBy?: { field: string; ascending: boolean };
  filter?: string;
  top?: number;
  skip?: number;
}

export type TCellResolvedValue = string | number | boolean | null | undefined | Record<string, unknown> | TCellResolvedValue[];

export interface ITableRendererProps {
  item: Record<string, unknown>;
  column: ITableColumnConfig;
  resolvedValue: TCellResolvedValue;
}

export type TTableRenderer = (props: ITableRendererProps) => ReactNode;
