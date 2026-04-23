import type { IFieldMetadata } from '../../../../../services/shared/types';
import type { ITableColumnConfig, ITableConfig, InternalFieldType } from '../types';
import { getInternalFieldType } from './fieldTypeMapper';
import {
  DEFAULT_FALLBACK_VALUE,
  DEFAULT_EMPTY_MESSAGE,
  ALIGN_BY_FIELD_TYPE,
  SORTABLE_BY_DEFAULT,
  EXPANDABLE_FIELD_TYPES,
  DEFAULT_EXPAND_DISPLAY_FIELD,
} from '../constants/tableDefaults';

function defaultAlign(fieldType: InternalFieldType): 'left' | 'center' | 'right' {
  return ALIGN_BY_FIELD_TYPE[fieldType] ?? 'left';
}

function defaultSortable(fieldType: InternalFieldType): boolean {
  return SORTABLE_BY_DEFAULT.indexOf(fieldType) !== -1;
}

export function normalizeColumnConfig(
  column: Partial<ITableColumnConfig>,
  fieldMetadata?: IFieldMetadata
): ITableColumnConfig {
  const internalName = column.internalName ?? column.id ?? '';
  const meta = fieldMetadata;
  const fieldType: InternalFieldType = column.fieldType ?? (meta ? getInternalFieldType(meta.TypeAsString, meta.MappedType) : 'unknown');

  const needsExpand = EXPANDABLE_FIELD_TYPES.indexOf(fieldType) !== -1;
  const expandConfig = column.expandConfig ?? (needsExpand
    ? { displayField: meta?.LookupField ?? DEFAULT_EXPAND_DISPLAY_FIELD }
    : undefined);

  const nonSortableMulti =
    fieldType === 'lookupMulti' || fieldType === 'userMulti' || fieldType === 'multiChoice';
  const nonSortableNote = fieldType === 'note';
  const sortable = nonSortableMulti || nonSortableNote ? false : (column.sortable ?? defaultSortable(fieldType));

  return {
    id: column.id ?? internalName,
    internalName,
    label: column.label ?? meta?.Title ?? internalName,
    visible: column.visible !== false,
    sortable,
    fieldType,
    width: column.width,
    minWidth: column.minWidth,
    maxWidth: column.maxWidth,
    align: column.align ?? defaultAlign(fieldType),
    fallbackValue: column.fallbackValue ?? DEFAULT_FALLBACK_VALUE,
    renderMode: column.renderMode,
    expandConfig,
    formatting: column.formatting,
    source: column.source,
  };
}

export function normalizeTableConfig(
  config: Partial<ITableConfig>,
  fieldsMetadata: IFieldMetadata[] = []
): ITableConfig {
  const byName = new Map(fieldsMetadata.map((f) => [f.InternalName, f]));
  const columns: ITableColumnConfig[] = (config.columns ?? []).map((c) =>
    normalizeColumnConfig(c, byName.get(c.internalName ?? c.id ?? ''))
  );

  let defaultSort = config.defaultSort;
  if (defaultSort?.field) {
    const root = defaultSort.field.split('/')[0];
    let colForSort: ITableColumnConfig | undefined;
    for (let i = 0; i < columns.length; i++) {
      if (columns[i].internalName === root) {
        colForSort = columns[i];
        break;
      }
    }
    if (colForSort && !colForSort.sortable) {
      defaultSort = undefined;
    }
  }

  return {
    enabled: config.enabled !== false,
    columns,
    sortable: config.sortable !== false,
    defaultSort,
    allowColumnToggle: config.allowColumnToggle,
    allowColumnReorder: config.allowColumnReorder,
    stickyHeader: config.stickyHeader,
    dense: config.dense,
    emptyMessage: config.emptyMessage ?? DEFAULT_EMPTY_MESSAGE,
  };
}
