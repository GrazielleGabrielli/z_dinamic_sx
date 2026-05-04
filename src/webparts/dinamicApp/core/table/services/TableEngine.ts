import type { IFieldMetadata } from '../../../../../services/shared/types';
import type {
  ITableConfig,
  ITableColumnConfig,
  ISortConfig,
  ITableDataRequest,
  TCellResolvedValue,
} from '../types';
import { normalizeTableConfig } from '../utils/tableConfigNormalizer';
import { buildSelect, buildExpand, buildSelectExpand } from '../utils/selectExpandBuilder';
import { columnODataPath } from '../utils/columnODataPath';
import { buildOrderBy } from '../utils/sortBuilder';
import {
  resolveRawCellValue,
  resolveDisplayCellValue,
  resolveFallbackValue,
} from '../utils/valueResolver';
import { getRenderer } from '../renderers/registry';

export class TableEngine {
  private _config: ITableConfig | null = null;

  normalizeTableConfig(config: Partial<ITableConfig>, fieldsMetadata: IFieldMetadata[] = []): ITableConfig {
    this._config = normalizeTableConfig(config, fieldsMetadata);
    return this._config;
  }

  getConfig(): ITableConfig | null {
    return this._config;
  }

  getVisibleColumns(config?: ITableConfig): ITableColumnConfig[] {
    const c = config ?? this._config;
    if (!c) return [];
    return c.columns.filter((col) => col.visible);
  }

  buildSelectExpand(columns?: ITableColumnConfig[]): { select: string[]; expand: string[] } {
    const cols = columns ?? this.getVisibleColumns();
    return buildSelectExpand(cols);
  }

  buildSelect(columns?: ITableColumnConfig[]): string[] {
    const cols = columns ?? this.getVisibleColumns();
    return buildSelect(cols);
  }

  buildExpand(columns?: ITableColumnConfig[]): string[] {
    const cols = columns ?? this.getVisibleColumns();
    return buildExpand(cols);
  }

  buildSort(sortConfig: ISortConfig | null | undefined): { field: string; ascending: boolean } | undefined {
    if (!sortConfig?.field) return undefined;
    const cols = this._config?.columns ?? [];
    const field = sortConfig.field;
    let blocked = false;
    for (let i = 0; i < cols.length; i++) {
      const c = cols[i];
      const path = columnODataPath(c);
      const prefix = c.internalName + '/';
      if (field === path || field === c.internalName || field.indexOf(prefix) === 0) {
        if (!c.sortable) blocked = true;
        break;
      }
    }
    if (blocked) return undefined;
    return buildOrderBy(sortConfig);
  }

  buildDataRequest(options: {
    sortConfig?: ISortConfig | null;
    filter?: string;
    top?: number;
    skip?: number;
  }): ITableDataRequest {
    const columns = this.getVisibleColumns();
    const { select, expand } = this.buildSelectExpand(columns);
    const orderBy = this.buildSort(options.sortConfig);
    return {
      select,
      expand: expand.length ? expand : undefined,
      orderBy,
      filter: options.filter,
      top: options.top,
      skip: options.skip,
    };
  }

  resolveCellValue(item: Record<string, unknown>, column: ITableColumnConfig): TCellResolvedValue {
    return resolveRawCellValue(item, column);
  }

  resolveDisplayValue(item: Record<string, unknown>, column: ITableColumnConfig): string {
    return resolveDisplayCellValue(item, column);
  }

  getRenderer(column: ITableColumnConfig) {
    return getRenderer(column.fieldType);
  }

  getFallbackValue(column: ITableColumnConfig): string {
    return resolveFallbackValue(column);
  }
}
