import type { InternalFieldType, TTableColumnAlign } from '../types';

export const DEFAULT_FALLBACK_VALUE = '—';

export const DEFAULT_EMPTY_MESSAGE = 'Nenhum item encontrado.';

export const DEFAULT_SEPARATOR_MULTI = ', ';

export const DEFAULT_BOOLEAN_TRUE = 'Sim';
export const DEFAULT_BOOLEAN_FALSE = 'Não';

export const ALIGN_BY_FIELD_TYPE: Partial<Record<InternalFieldType, TTableColumnAlign>> = {
  number: 'right',
  currency: 'right',
  boolean: 'center',
  date: 'center',
};

export const SORTABLE_BY_DEFAULT: InternalFieldType[] = [
  'text',
  'number',
  'currency',
  'date',
  'boolean',
  'choice',
  'lookup',
  'user',
  'url',
];

export const EXPANDABLE_FIELD_TYPES: InternalFieldType[] = [
  'lookup',
  'lookupMulti',
  'user',
  'userMulti',
];

export const DEFAULT_EXPAND_DISPLAY_FIELD = 'Title';
