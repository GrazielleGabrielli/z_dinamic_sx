import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  IconButton,
  TooltipHost,
  Dropdown,
  IDropdownOption,
} from '@fluentui/react';
import { IDynamicViewConfig, IListViewColumnConfig } from '../../core/config/types';
import { buildListQuery } from '../../core/listView';
import { buildDynamicContext, parseQueryString } from '../../core/dynamicTokens';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import { ItemsService, UsersService, FieldsService } from '../../../../services';
import type { IFieldMetadata } from '../../../../services';
import { readListItemId } from '../../../../services/items/listItemId';

const DEFAULT_COLUMN_COUNT = 10;

function columnConfigToIColumn(c: IListViewColumnConfig, index: number): IColumn {
  const base: IColumn = {
    key: c.field,
    name: c.label ?? c.field,
    fieldName: c.field,
    minWidth: c.width ?? 80,
  };
  const expandField = c.expandField;
  if (expandField) {
    base.onRender = (item: Record<string, unknown>) => {
      const val = item[c.field];
      const v = val && typeof val === 'object' && expandField in (val as object)
        ? (val as Record<string, unknown>)[expandField]
        : val;
      return v !== null && v !== undefined ? String(v) : '';
    };
  }
  return base;
}

function fieldToColumnConfig(f: IFieldMetadata): IListViewColumnConfig {
  const needsExpand = ['lookup', 'lookupmulti', 'user', 'usermulti'].indexOf(f.MappedType) !== -1;
  return {
    field: f.InternalName,
    label: f.Title,
    ...(needsExpand && { expandField: f.LookupField || 'Title' }),
  };
}

export const ListView: React.FC<{ config: IDynamicViewConfig }> = ({ config }) => {
  const { listView, pagination, dataSource } = config;
  const listTitle = dataSource.title;

  const [defaultColumns, setDefaultColumns] = useState<IListViewColumnConfig[]>([]);
  const [defaultColumnsLoading, setDefaultColumnsLoading] = useState(false);
  const [enrichedConfigColumns, setEnrichedConfigColumns] = useState<IListViewColumnConfig[] | undefined>(undefined);
  const [listFieldMetadata, setListFieldMetadata] = useState<IFieldMetadata[] | undefined>(undefined);
  const [items, setItems] = useState<Record<string, unknown>[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | undefined>(undefined);
  const [paging, setPaging] = useState<{
    pageIndex: number;
    forwardPivots: number[];
    resetKey: string | null;
  }>({ pageIndex: 0, forwardPivots: [], resetKey: null });
  const [hasNext, setHasNext] = useState(false);
  const [selectedViewModeId, setSelectedViewModeId] = useState<string>(
    () => listView?.activeViewModeId ?? listView?.viewModes?.[0]?.id ?? 'all'
  );
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

  const hasConfigColumns = (listView.columns?.length ?? 0) > 0;

  useEffect(() => {
    if (!listTitle.trim()) return;
    setListFieldMetadata(undefined);
    const expandableTypes = ['lookup', 'lookupmulti', 'user', 'usermulti'];
    if (!hasConfigColumns) {
      setEnrichedConfigColumns(undefined);
      setDefaultColumnsLoading(true);
      fieldsService
        .getVisibleFields(listTitle.trim())
        .then((fields) => {
          setListFieldMetadata(fields);
          setDefaultColumns(fields.slice(0, DEFAULT_COLUMN_COUNT).map(fieldToColumnConfig));
        })
        .then(() => setDefaultColumnsLoading(false), () => setDefaultColumnsLoading(false));
      return;
    }
    setEnrichedConfigColumns(undefined);
    fieldsService
      .getVisibleFields(listTitle.trim())
      .then((fields) => {
        setListFieldMetadata(fields);
        const byName = new Map(fields.map((f) => [f.InternalName, f]));
        const enriched = (listView.columns ?? []).map((col) => {
          const meta = byName.get(col.field);
          const needsExpand = meta && expandableTypes.indexOf(meta.MappedType) !== -1;
          return needsExpand && !col.expandField
            ? { ...col, expandField: meta.LookupField || 'Title' }
            : col;
        });
        setEnrichedConfigColumns(enriched);
      })
      .catch(() => setEnrichedConfigColumns(listView.columns ?? []));
  }, [listTitle, hasConfigColumns, listView.columns]);

  const effectiveColumns = useMemo(() => {
    if (hasConfigColumns) {
      return enrichedConfigColumns ?? [];
    }
    return defaultColumns;
  }, [hasConfigColumns, listView.columns, enrichedConfigColumns, defaultColumns]);

  const effectiveListView = useMemo(
    () => ({
      ...listView,
      columns: effectiveColumns,
      activeViewModeId: selectedViewModeId,
    }),
    [listView, effectiveColumns, selectedViewModeId]
  );

  const queryOptions = useMemo(
    () => buildListQuery(effectiveListView, { dynamicContext, fieldsMetadata: listFieldMetadata }),
    [effectiveListView, dynamicContext, listFieldMetadata]
  );

  const pagingResetKey = useMemo(
    () =>
      [
        listTitle.trim(),
        String(pagination.enabled),
        String(pagination.pageSize),
        selectedViewModeId,
        queryOptions.filter ?? '',
        queryOptions.orderBy?.field ?? '',
        String(queryOptions.orderBy?.ascending ?? ''),
        (queryOptions.select ?? []).join(','),
        (queryOptions.expand ?? []).join(','),
        String(effectiveColumns.length),
        String(listFieldMetadata?.length ?? ''),
      ].join('|'),
    [
      listTitle,
      pagination.enabled,
      pagination.pageSize,
      selectedViewModeId,
      queryOptions.filter,
      queryOptions.orderBy?.field,
      queryOptions.orderBy?.ascending,
      queryOptions.select,
      queryOptions.expand,
      effectiveColumns.length,
      listFieldMetadata?.length,
    ]
  );

  useEffect(() => {
    if (!listTitle.trim()) {
      setItems([]);
      setLoading(false);
      return;
    }
    if (effectiveColumns.length === 0) {
      setLoading(hasConfigColumns ? true : defaultColumnsLoading);
      return;
    }
    if (listFieldMetadata === undefined) {
      setLoading(true);
      return;
    }

    if (paging.resetKey !== pagingResetKey) {
      setPaging({ pageIndex: 0, forwardPivots: [], resetKey: pagingResetKey });
      return;
    }

    setLoading(true);
    setError(undefined);

    const pageSize = pagination.enabled ? pagination.pageSize : 100;
    const options = {
      select: queryOptions.select,
      expand: queryOptions.expand,
      filter: queryOptions.filter,
      orderBy: queryOptions.orderBy,
      fieldMetadata: listFieldMetadata,
    };

    const afterLastItemId =
      paging.pageIndex === 0 ? undefined : paging.forwardPivots[paging.pageIndex - 1];

    let cancelled = false;
    itemsService
      .getPagedItems<Record<string, unknown>>(listTitle.trim(), options, pageSize, afterLastItemId)
      .then((result) => {
        if (cancelled) return;
        setItems(result.items);
        setHasNext(result.hasNext);
        setLoading(false);
      })
      .catch((err: Error) => {
        if (cancelled) return;
        setError(err.message);
        setItems([]);
        setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [
    itemsService,
    listTitle,
    pagination.enabled,
    pagination.pageSize,
    defaultColumnsLoading,
    hasConfigColumns,
    effectiveColumns.length,
    listFieldMetadata,
    pagingResetKey,
    paging,
    queryOptions.select,
    queryOptions.expand,
    queryOptions.filter,
    queryOptions.orderBy?.field,
    queryOptions.orderBy?.ascending,
  ]);

  const columns: IColumn[] = useMemo(
    () => effectiveColumns.map((c, i) => columnConfigToIColumn(c, i)),
    [effectiveColumns]
  );

  const viewModes = listView?.viewModes ?? [];
  const viewModeOptions: IDropdownOption[] = viewModes.map((m) => ({ key: m.id, text: m.label }));

  const currentPage = paging.pageIndex + 1;

  if (error) {
    return (
      <MessageBar messageBarType={MessageBarType.error} isMultiline={false} styles={{ root: { marginTop: 12 } }}>
        {error}
      </MessageBar>
    );
  }

  if (loading && items.length === 0) {
    return (
      <Stack horizontalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { padding: '32px 0' } }}>
        <Spinner size={SpinnerSize.medium} />
        <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
          Carregando itens...
        </Text>
      </Stack>
    );
  }

  if (items.length === 0) {
    return (
      <Stack
        tokens={{ childrenGap: 8 }}
        styles={{ root: { marginTop: 16, padding: 24, background: '#faf9f8', borderRadius: 8, border: '1px solid #edebe9' } }}
      >
        <Text variant="medium" styles={{ root: { color: '#605e5c' } }}>
          Nenhum item encontrado.
        </Text>
        <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
          A lista está vazia ou os filtros não retornaram resultados.
        </Text>
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 8 } }}>
      {viewModeOptions.length > 0 && (
        <Dropdown
          label="Visualização"
          options={viewModeOptions}
          selectedKey={selectedViewModeId}
          onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
            if (opt) setSelectedViewModeId(String(opt.key));
          }}
          styles={{ root: { maxWidth: 220 } }}
        />
      )}
      <DetailsList
        items={items}
        columns={columns}
        layoutMode={DetailsListLayoutMode.justified}
        selectionMode={SelectionMode.none}
        compact={false}
        isHeaderVisible={true}
      />

      {pagination.enabled && (
        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          tokens={{ childrenGap: 8 }}
          styles={{ root: { padding: '8px 0', borderTop: '1px solid #edebe9' } }}
        >
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Página {currentPage}
            {items.length > 0 && ` · ${items.length} itens nesta página`}
          </Text>
          <Stack horizontal tokens={{ childrenGap: 4 }}>
            <TooltipHost content="Página anterior">
              <IconButton
                iconProps={{ iconName: 'ChevronLeft' }}
                disabled={paging.pageIndex === 0}
                onClick={() => {
                  setPaging((prev) =>
                    prev.resetKey !== pagingResetKey || prev.pageIndex <= 0
                      ? prev
                      : { ...prev, pageIndex: prev.pageIndex - 1 }
                  );
                }}
              />
            </TooltipHost>
            <TooltipHost content="Próxima página">
              <IconButton
                iconProps={{ iconName: 'ChevronRight' }}
                disabled={!hasNext}
                onClick={() => {
                  const last = readListItemId(items[items.length - 1]);
                  if (last === undefined) return;
                  setPaging((prev) => {
                    if (prev.resetKey !== pagingResetKey) return prev;
                    const pivots = prev.forwardPivots.slice(0, prev.pageIndex);
                    pivots[prev.pageIndex] = last;
                    return { ...prev, pageIndex: prev.pageIndex + 1, forwardPivots: pivots };
                  });
                }}
              />
            </TooltipHost>
          </Stack>
        </Stack>
      )}
    </Stack>
  );
};
