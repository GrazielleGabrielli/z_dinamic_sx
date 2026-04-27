import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import { ActionButton, MessageBar, MessageBarType, Spinner, Stack, Text } from '@fluentui/react';
import type { IFieldMetadata } from '../../../../services';
import type {
  IDynamicViewConfig,
  IListViewColumnConfig,
  IListViewFilterConfig,
  IProjectManagementColumnConfig,
} from '../../core/config/types';
import { buildListFilter, getActiveViewModeFilters, isNoteFieldPath } from '../../core/listView';
import { buildDynamicContext, parseQueryString } from '../../core/dynamicTokens';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import { FieldsService, ItemsService, UsersService } from '../../../../services';

export interface IProjectManagementViewProps {
  config: IDynamicViewConfig;
  dashboardListFilters?: IListViewFilterConfig[];
  onItemUpdated?: () => void;
  onEditTableColumns?: () => void;
}

function valueToText(value: unknown, expandField?: string): string {
  if (value === null || value === undefined) return '';
  if (Array.isArray(value)) return value.map((x) => valueToText(x, expandField)).filter(Boolean).join(', ');
  if (typeof value === 'object') {
    const obj = value as Record<string, unknown>;
    if (expandField && obj[expandField] !== undefined) return valueToText(obj[expandField]);
    if (obj.Title !== undefined) return valueToText(obj.Title);
    if (obj.Label !== undefined) return valueToText(obj.Label);
  }
  return String(value);
}

function buildSelectExpand(
  columns: IListViewColumnConfig[],
  ruleFields: string[]
): { select: string[]; expand: string[] } {
  const select: string[] = ['Id'];
  const expand: string[] = [];
  const addSimple = (field: string): void => {
    if (select.indexOf(field) === -1) select.push(field);
  };
  const addExpand = (field: string, expandField: string): void => {
    if (expand.indexOf(field) === -1) expand.push(field);
    if (select.indexOf(`${field}/Id`) === -1) select.push(`${field}/Id`);
    if (select.indexOf(`${field}/${expandField}`) === -1) select.push(`${field}/${expandField}`);
  };
  for (let i = 0; i < columns.length; i++) {
    const col = columns[i];
    if (!col.field) continue;
    if (col.expandField) addExpand(col.field, col.expandField);
    else addSimple(col.field);
  }
  for (let i = 0; i < ruleFields.length; i++) addSimple(ruleFields[i]);
  if (select.indexOf('Title') === -1) select.push('Title');
  return { select, expand };
}

function itemMatchesColumn(item: Record<string, unknown>, column: IProjectManagementColumnConfig): boolean {
  if (!column.rules || column.rules.length === 0) return false;
  for (let i = 0; i < column.rules.length; i++) {
    const rule = column.rules[i];
    const current = valueToText(item[rule.field]).trim();
    if (current !== rule.value.trim()) return false;
  }
  return true;
}

function coerceRuleValue(value: string, fieldMeta: IFieldMetadata | undefined): unknown {
  if (!fieldMeta) return value;
  if (fieldMeta.MappedType === 'number' || fieldMeta.MappedType === 'currency') {
    const n = Number(value);
    return isNaN(n) ? value : n;
  }
  if (fieldMeta.MappedType === 'boolean') {
    const v = value.trim().toLowerCase();
    if (v === 'true' || v === '1' || v === 'sim') return true;
    if (v === 'false' || v === '0' || v === 'nao' || v === 'não') return false;
  }
  return value;
}

export const ProjectManagementView: React.FC<IProjectManagementViewProps> = ({
  config,
  dashboardListFilters,
  onItemUpdated,
  onEditTableColumns,
}) => {
  const { dataSource, listView } = config;
  const pm = config.projectManagement;
  const listTitle = dataSource.title;
  const listWeb = dataSource.webServerRelativeUrl?.trim() || undefined;
  const columns = pm?.columns ?? [];
  const titleColumn = listView.columns[0] ?? { field: 'Title', label: 'Title' };
  const detailColumns = listView.columns.filter((c) => c.field !== titleColumn.field).slice(0, 3);
  const ruleFields = useMemo(() => {
    const next: string[] = [];
    for (let i = 0; i < columns.length; i++) {
      const rules = columns[i].rules ?? [];
      for (let j = 0; j < rules.length; j++) {
        const field = rules[j].field.trim();
        if (field && next.indexOf(field) === -1) next.push(field);
      }
    }
    return next;
  }, [columns]);

  const [items, setItems] = useState<Record<string, unknown>[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | undefined>(undefined);
  const [fieldMetadata, setFieldMetadata] = useState<IFieldMetadata[]>([]);
  const [draggingId, setDraggingId] = useState<number | null>(null);
  const [updatingId, setUpdatingId] = useState<number | null>(null);
  const [dynamicContext, setDynamicContext] = useState<IDynamicContext | undefined>(undefined);

  const itemsService = useMemo(() => new ItemsService(), []);
  const fieldsService = useMemo(() => new FieldsService(), []);

  useEffect(() => {
    const usersService = new UsersService();
    usersService
      .getCurrentUser()
      .then((user) => {
        setDynamicContext(
          buildDynamicContext({
            currentUser: { id: user.Id, title: user.Title, name: user.Title, email: user.Email, loginName: user.LoginName },
            query: typeof window !== 'undefined' && window.location ? parseQueryString(window.location.search) : undefined,
            now: new Date(),
          })
        );
      })
      .catch(() => setDynamicContext(buildDynamicContext({ now: new Date() })));
  }, []);

  useEffect(() => {
    if (!listTitle.trim()) return;
    setLoading(true);
    setError(undefined);
    fieldsService
      .getVisibleFields(listTitle, listWeb)
      .then((fieldMetadata) => {
        setFieldMetadata(fieldMetadata);
        const selectExpand = buildSelectExpand(listView.columns, ruleFields);
        const listViewWithMode = { ...listView, activeViewModeId: listView.activeViewModeId ?? listView.viewModes?.[0]?.id ?? 'all' };
        const viewModeFilters = getActiveViewModeFilters(listViewWithMode);
        const viewModeFilterStr = buildListFilter(viewModeFilters, { dynamicContext, fieldsMetadata: fieldMetadata });
        const dashboardFilterStr =
          dashboardListFilters && dashboardListFilters.length > 0
            ? buildListFilter(dashboardListFilters, { dynamicContext, fieldsMetadata: fieldMetadata })
            : undefined;
        const filterParts = [viewModeFilterStr, dashboardFilterStr].filter(Boolean);
        const filter = filterParts.length > 0 ? filterParts.join(' and ') : undefined;
        const orderBy =
          listView.sort?.field && !isNoteFieldPath(listView.sort.field, fieldMetadata)
            ? { field: listView.sort.field, ascending: listView.sort.ascending }
            : undefined;
        return itemsService.getItems<Record<string, unknown>>(listTitle, {
          select: selectExpand.select,
          expand: selectExpand.expand.length > 0 ? selectExpand.expand : undefined,
          filter,
          orderBy,
          top: 500,
          fieldMetadata,
          ...(listWeb ? { webServerRelativeUrl: listWeb } : {}),
        });
      })
      .then((result) => {
        setItems(result);
        setLoading(false);
      })
      .catch((e: Error) => {
        setError(e.message);
        setItems([]);
        setLoading(false);
      });
  }, [dashboardListFilters, dynamicContext, fieldsService, itemsService, listTitle, listWeb, listView, ruleFields]);

  const grouped = useMemo(() => {
    const groups: Record<string, Record<string, unknown>[]> = {};
    for (let i = 0; i < columns.length; i++) groups[columns[i].id] = [];
    const ungrouped: Record<string, unknown>[] = [];
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      let matched = false;
      for (let j = 0; j < columns.length; j++) {
        if (itemMatchesColumn(item, columns[j])) {
          groups[columns[j].id].push(item);
          matched = true;
          break;
        }
      }
      if (!matched) ungrouped.push(item);
    }
    return { groups, ungrouped };
  }, [columns, items]);

  const handleDrop = async (target: IProjectManagementColumnConfig): Promise<void> => {
    if (!draggingId) return;
    const updateValues: Record<string, unknown> = {};
    for (let i = 0; i < target.rules.length; i++) {
      const rule = target.rules[i];
      if (!rule.field.trim()) continue;
      let meta: IFieldMetadata | undefined;
      for (let j = 0; j < fieldMetadata.length; j++) {
        if (fieldMetadata[j].InternalName === rule.field) {
          meta = fieldMetadata[j];
          break;
        }
      }
      updateValues[rule.field] = coerceRuleValue(rule.value, meta);
    }
    if (Object.keys(updateValues).length === 0) return;
    setUpdatingId(draggingId);
    setError(undefined);
    try {
      await itemsService.updateItem(listTitle, draggingId, updateValues, listWeb);
      setItems((prev) =>
        prev.map((item) =>
          Number(item.Id) === draggingId ? { ...item, ...updateValues } : item
        )
      );
      if (onItemUpdated) onItemUpdated();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setDraggingId(null);
      setUpdatingId(null);
    }
  };

  if (columns.length === 0) {
    return (
      <Stack tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 8 } }}>
        {onEditTableColumns !== undefined ? (
          <Stack horizontal horizontalAlign="end" styles={{ root: { width: '100%' } }}>
            <ActionButton
              iconProps={{ iconName: 'ColumnOptions' }}
              onClick={onEditTableColumns}
              styles={{ root: { height: 30, color: '#0078d4' } }}
            >
              Editar quadro
            </ActionButton>
          </Stack>
        ) : null}
        <MessageBar messageBarType={MessageBarType.warning}>Adicione pelo menos uma coluna no quadro.</MessageBar>
      </Stack>
    );
  }
  if (loading) {
    return (
      <Stack horizontalAlign="center" styles={{ root: { padding: 32 } }}>
        <Spinner label="Carregando quadro..." />
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 8 } }}>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {onEditTableColumns !== undefined ? (
        <Stack horizontal horizontalAlign="end" styles={{ root: { width: '100%' } }}>
          <ActionButton
            iconProps={{ iconName: 'ColumnOptions' }}
            onClick={onEditTableColumns}
            styles={{ root: { height: 30, color: '#0078d4' } }}
          >
            Editar quadro
          </ActionButton>
        </Stack>
      ) : null}
      <Stack horizontal wrap tokens={{ childrenGap: 12 }} verticalAlign="start" styles={{ root: { overflowX: 'auto', paddingBottom: 8 } }}>
        {columns.map((col) => (
          <Stack
            key={col.id}
            tokens={{ childrenGap: 10 }}
            styles={{
              root: {
                width: 290,
                minHeight: 320,
                padding: 12,
                borderRadius: 8,
                border: '1px solid #edebe9',
                background: '#faf9f8',
                boxSizing: 'border-box',
              },
            }}
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => {
              e.preventDefault();
              void handleDrop(col);
            }}
          >
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>{col.title}</Text>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {(col.rules ?? []).length} regra(s)
              </Text>
            </Stack>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              {grouped.groups[col.id]?.length ?? 0} item(ns)
            </Text>
            {(col.rules ?? []).length > 0 && (
              <Stack tokens={{ childrenGap: 2 }}>
                {col.rules.map((rule) => (
                  <Text key={rule.id} variant="small" styles={{ root: { color: '#605e5c' } }}>
                    {rule.field} = {rule.value || '""'}
                  </Text>
                ))}
              </Stack>
            )}
            {(grouped.groups[col.id] ?? []).map((item) => {
              const itemId = Number(item.Id);
              const disabled = updatingId === itemId;
              return (
                <div
                  key={itemId}
                  draggable={!disabled}
                  onDragStart={(e) => {
                    setDraggingId(itemId);
                    e.dataTransfer.effectAllowed = 'move';
                    e.dataTransfer.setData('text/plain', String(itemId));
                  }}
                  onDragEnd={() => setDraggingId(null)}
                  style={{
                    padding: 12,
                    borderRadius: 8,
                    border: draggingId === itemId ? '2px solid #0078d4' : '1px solid #edebe9',
                    background: '#fff',
                    cursor: disabled ? 'progress' : 'grab',
                    opacity: disabled ? 0.6 : 1,
                    boxShadow: '0 1px 3px rgba(0,0,0,0.06)',
                  }}
                >
                  <Text variant="medium" styles={{ root: { fontWeight: 600, marginBottom: 6 } }}>
                    {valueToText(item[titleColumn.field], titleColumn.expandField) || `Item ${itemId}`}
                  </Text>
                  {detailColumns.map((detail) => (
                    <Text key={detail.field} variant="small" styles={{ root: { color: '#605e5c', display: 'block', marginTop: 4 } }}>
                      {(detail.label ?? detail.field)}: {valueToText(item[detail.field], detail.expandField) || '—'}
                    </Text>
                  ))}
                </div>
              );
            })}
          </Stack>
        ))}
        {grouped.ungrouped.length > 0 && (
          <Stack
            tokens={{ childrenGap: 10 }}
            styles={{
              root: {
                width: 290,
                minHeight: 320,
                padding: 12,
                borderRadius: 8,
                border: '1px dashed #c8c6c4',
                background: '#fff',
                boxSizing: 'border-box',
              },
            }}
          >
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>Sem coluna</Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              O item ainda não atende às regras de nenhuma coluna configurada.
            </Text>
            {grouped.ungrouped.map((item) => {
              const itemId = Number(item.Id);
              const disabled = updatingId === itemId;
              return (
                <div
                  key={itemId}
                  draggable={!disabled}
                  onDragStart={(e) => {
                    setDraggingId(itemId);
                    e.dataTransfer.effectAllowed = 'move';
                    e.dataTransfer.setData('text/plain', String(itemId));
                  }}
                  onDragEnd={() => setDraggingId(null)}
                  style={{
                    padding: 12,
                    borderRadius: 8,
                    border: draggingId === itemId ? '2px solid #0078d4' : '1px solid #edebe9',
                    background: '#fff',
                    cursor: disabled ? 'progress' : 'grab',
                    opacity: disabled ? 0.6 : 1,
                    boxShadow: '0 1px 3px rgba(0,0,0,0.06)',
                  }}
                >
                  <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                    {valueToText(item[titleColumn.field], titleColumn.expandField) || `Item ${itemId}`}
                  </Text>
                  {detailColumns.map((detail) => (
                    <Text key={detail.field} variant="small" styles={{ root: { color: '#605e5c', display: 'block', marginTop: 4 } }}>
                      {(detail.label ?? detail.field)}: {valueToText(item[detail.field], detail.expandField) || '—'}
                    </Text>
                  ))}
                </div>
              );
            })}
          </Stack>
        )}
      </Stack>
    </Stack>
  );
};
