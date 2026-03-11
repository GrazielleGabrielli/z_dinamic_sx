import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  Checkbox,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  Separator,
  ChoiceGroup,
  IChoiceGroupOption,
  IconButton,
} from '@fluentui/react';
import { FieldsService } from '../../../../services';
import type { IFieldMetadata } from '../../../../services';
import type {
  IListViewConfig,
  IListViewColumnConfig,
  IListViewModeConfig,
  IListViewFilterConfig,
  IPaginationConfig,
  TPaginationLayout,
  TFilterOperator,
} from '../../core/config/types';

interface ITableColumnsEditorPanelProps {
  isOpen: boolean;
  listTitle: string;
  listView: IListViewConfig;
  pagination: IPaginationConfig;
  onSave: (listView: IListViewConfig, pagination: IPaginationConfig) => void;
  onDismiss: () => void;
}

interface IFieldOption {
  meta: IFieldMetadata;
  selected: boolean;
  label: string;
  expandField: string;
}

const EXPANDABLE = ['lookup', 'lookupmulti', 'user', 'usermulti'];

const SIMPLE_FIELD_TYPES = ['text', 'multiline', 'number', 'currency', 'boolean', 'choice', 'multichoice', 'datetime', 'url'];

const USER_EXPAND_FIELDS: IDropdownOption[] = [
  { key: 'Id', text: 'Id' },
  { key: 'Title', text: 'Title' },
  { key: 'EMail', text: 'EMail' },
  { key: 'LoginName', text: 'LoginName' },
];

function toFieldOption(meta: IFieldMetadata, existing?: IListViewColumnConfig): IFieldOption {
  const selected = existing !== undefined;
  const needsExpand = EXPANDABLE.indexOf(meta.MappedType) !== -1;
  const expandField = needsExpand
    ? (existing?.expandField ?? meta.LookupField ?? 'Title')
    : '';
  return {
    meta,
    selected,
    label: existing?.label ?? meta.Title,
    expandField,
  };
}

function buildOptions(
  fields: IFieldMetadata[],
  currentColumns: IListViewColumnConfig[]
): IFieldOption[] {
  const byName = new Map(fields.map((f) => [f.InternalName, f]));
  const selectedSet = new Set(currentColumns.map((c) => c.field));
  const ordered: IFieldOption[] = [];
  currentColumns.forEach((c) => {
    const meta = byName.get(c.field);
    if (meta) ordered.push(toFieldOption(meta, c));
  });
  fields.forEach((f) => {
    if (!selectedSet.has(f.InternalName)) ordered.push(toFieldOption(f, undefined));
  });
  return ordered;
}

function buildExpandOptionsFromLookupList(fields: IFieldMetadata[]): IDropdownOption[] {
  const simple = fields.filter(
    (f) =>
      SIMPLE_FIELD_TYPES.indexOf(f.MappedType) !== -1 &&
      f.InternalName !== 'Id' &&
      f.InternalName !== 'Title'
  );
  const options: IDropdownOption[] = [
    { key: 'Id', text: 'Id' },
    { key: 'Title', text: 'Title' },
  ];
  simple.forEach((f) => options.push({ key: f.InternalName, text: `${f.Title} (${f.InternalName})` }));
  return options;
}

const PAGE_SIZE_OPTIONS = [5, 10, 20, 50, 100];

const PAGINATION_LAYOUT_OPTIONS: IChoiceGroupOption[] = [
  { key: 'buttons', text: 'Botões (Anterior / Próxima)' },
  { key: 'numbered', text: 'Numerada (Página X + botões)' },
  { key: 'compact', text: 'Compacta (linha única)' },
  { key: 'paged', text: 'Páginas (1 … 4 5 6 …)' },
];

const VIEW_MODE_OPERATORS: IDropdownOption[] = [
  { key: 'eq', text: 'Igual a' },
  { key: 'ne', text: 'Diferente de' },
  { key: 'contains', text: 'Contém' },
  { key: 'gt', text: 'Maior que' },
  { key: 'ge', text: 'Maior ou igual' },
  { key: 'lt', text: 'Menor que' },
  { key: 'le', text: 'Menor ou igual' },
];

function viewModeFilterSummary(filters: IListViewFilterConfig[]): string {
  if (!filters || filters.length === 0) return 'Sem filtros';
  return filters.map((f) => `${f.field} ${f.operator} "${f.value}"`).join(' e ');
}

export const TableColumnsEditorPanel: React.FC<ITableColumnsEditorPanelProps> = ({
  isOpen,
  listTitle,
  listView,
  pagination,
  onSave,
  onDismiss,
}) => {
  const [loading, setLoading] = useState(false);
  const [options, setOptions] = useState<IFieldOption[]>([]);
  const [lookupListFields, setLookupListFields] = useState<Record<string, IFieldMetadata[]>>({});
  const [paginationEnabled, setPaginationEnabled] = useState(pagination.enabled);
  const [pageSize, setPageSize] = useState(pagination.pageSize);
  const [paginationLayout, setPaginationLayout] = useState<TPaginationLayout>(pagination.layout ?? 'buttons');
  const defaultViewModes: IListViewModeConfig[] = [
    { id: 'all', label: 'Todas', filters: [] },
    { id: 'mine', label: 'Minhas', filters: [{ field: 'Author/Id', operator: 'eq', value: '[Me]' }] },
  ];
  const [viewModes, setViewModes] = useState<IListViewModeConfig[]>(listView.viewModes?.length ? listView.viewModes : defaultViewModes);
  const [activeViewModeId, setActiveViewModeId] = useState<string>(listView.activeViewModeId ?? 'all');
  const [viewModeEditingId, setViewModeEditingId] = useState<string | null>(null);
  const [viewModeEditLabel, setViewModeEditLabel] = useState('');
  const [viewModeEditFilters, setViewModeEditFilters] = useState<IListViewFilterConfig[]>([]);

  const fieldsService = useMemo(() => new FieldsService(), []);

  useEffect(() => {
    if (!isOpen || !listTitle.trim()) return;
    setLoading(true);
    setLookupListFields({});
    fieldsService
      .getVisibleFields(listTitle.trim())
      .then((f) => {
        setOptions(buildOptions(f, listView.columns ?? []));
        const listIds = f
          .filter((x) => EXPANDABLE.indexOf(x.MappedType) !== -1 && x.LookupList)
          .map((x) => x.LookupList as string);
        const uniqueIds = listIds.filter((id, i) => listIds.indexOf(id) === i);
        return Promise.all(
          uniqueIds.map((id) =>
            fieldsService.getFields(id).then((fields) => ({ id, fields }))
          )
        );
      })
      .then((results) => {
        const next: Record<string, IFieldMetadata[]> = {};
        results.forEach(({ id, fields }) => { next[id] = fields; });
        setLookupListFields((prev) => ({ ...prev, ...next }));
      })
      .then(() => setLoading(false), () => setLoading(false));
  }, [isOpen, listTitle]);

  useEffect(() => {
    if (isOpen) {
      setPaginationEnabled(pagination.enabled);
      setPageSize(pagination.pageSize);
      setPaginationLayout(pagination.layout ?? 'buttons');
      setViewModes(listView.viewModes?.length ? listView.viewModes : defaultViewModes);
      setActiveViewModeId(listView.activeViewModeId ?? 'all');
    }
  }, [isOpen, pagination.enabled, pagination.pageSize, pagination.layout, listView.viewModes, listView.activeViewModeId]);


  const toggle = (internalName: string): void => {
    setOptions((prev) =>
      prev.map((o) => (o.meta.InternalName === internalName ? { ...o, selected: !o.selected } : o))
    );
  };

  const setLabel = (internalName: string, label: string): void => {
    setOptions((prev) =>
      prev.map((o) => (o.meta.InternalName === internalName ? { ...o, label } : o))
    );
  };

  const setExpandField = (internalName: string, expandField: string): void => {
    setOptions((prev) =>
      prev.map((o) => (o.meta.InternalName === internalName ? { ...o, expandField } : o))
    );
  };

  const getExpandFieldOptions = (meta: IFieldMetadata): IDropdownOption[] => {
    if (meta.MappedType === 'user' || meta.MappedType === 'usermulti') return USER_EXPAND_FIELDS;
    if (meta.LookupList && lookupListFields[meta.LookupList]) {
      return buildExpandOptionsFromLookupList(lookupListFields[meta.LookupList]);
    }
    return [{ key: 'Title', text: 'Title' }, { key: 'Id', text: 'Id' }];
  };

  const handleSave = (): void => {
    const columns: IListViewColumnConfig[] = options
      .filter((o) => o.selected)
      .map((o) => {
        const base = {
          field: o.meta.InternalName,
          label: o.label.trim() ? o.label : o.meta.Title,
        };
        if (EXPANDABLE.indexOf(o.meta.MappedType) !== -1) {
          return { ...base, expandField: o.expandField.trim() || 'Title' };
        }
        return base;
      });
    const nextPagination: IPaginationConfig = {
      ...pagination,
      enabled: paginationEnabled,
      pageSize,
      layout: paginationLayout,
      pageSizeOptions: pagination.pageSizeOptions?.length ? pagination.pageSizeOptions : PAGE_SIZE_OPTIONS,
    };
    onSave(
      { ...listView, columns, viewModes, activeViewModeId },
      nextPagination
    );
    onDismiss();
  };

  const viewModeDefaultOptions: IDropdownOption[] = viewModes.map((m) => ({ key: m.id, text: m.label }));
  const startViewModeAdd = (): void => {
    setViewModeEditLabel('Novo modo');
    setViewModeEditFilters([]);
    setViewModeEditingId(`custom_${Date.now()}`);
  };
  const startViewModeEdit = (m: IListViewModeConfig): void => {
    setViewModeEditingId(m.id);
    setViewModeEditLabel(m.label);
    setViewModeEditFilters(m.filters?.length ? m.filters.slice() : []);
  };
  const saveViewModeEdit = (): void => {
    if (viewModeEditingId === null) return;
    const next = viewModes.slice();
    let idx = -1;
    for (let i = 0; i < next.length; i++) { if (next[i].id === viewModeEditingId) { idx = i; break; } }
    const entry: IListViewModeConfig = { id: viewModeEditingId, label: viewModeEditLabel.trim() || viewModeEditingId, filters: viewModeEditFilters };
    if (idx >= 0) next[idx] = entry;
    else next.push(entry);
    setViewModes(next);
    setViewModeEditingId(null);
  };
  const removeViewMode = (id: string): void => {
    if (id === 'all' || id === 'mine') return;
    const next = viewModes.filter((m) => m.id !== id);
    setViewModes(next);
    if (activeViewModeId === id) setActiveViewModeId(next[0]?.id ?? 'all');
  };
  const addViewModeFilter = (): void => setViewModeEditFilters([...viewModeEditFilters, { field: '', operator: 'eq', value: '' }]);
  const removeViewModeFilter = (i: number): void => setViewModeEditFilters(viewModeEditFilters.filter((_, idx) => idx !== i));
  const updateViewModeFilter = (i: number, part: Partial<IListViewFilterConfig>): void => {
    const next = viewModeEditFilters.slice();
    next[i] = { ...next[i], ...part };
    setViewModeEditFilters(next);
  };

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Editar lista / tabela"
      closeButtonAriaLabel="Fechar"
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="Salvar" onClick={handleSave} />
          <DefaultButton text="Cancelar" onClick={onDismiss} />
        </Stack>
      )}
    >
      <div style={{ paddingTop: 16 }}>
        {loading ? (
          <Stack horizontalAlign="center" tokens={{ childrenGap: 12 }} style={{ padding: 24 }}>
            <Spinner size={SpinnerSize.medium} />
            <Text variant="small">Carregando campos...</Text>
          </Stack>
        ) : (
          <Stack tokens={{ childrenGap: 16 }}>
            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                Paginação
              </Text>
              <Checkbox
                label="Habilitar paginação"
                checked={paginationEnabled}
                onChange={(_, v) => setPaginationEnabled(!!v)}
              />
              {paginationEnabled && (
                <>
                  <Dropdown
                    label="Itens por página"
                    selectedKey={String(pageSize)}
                    options={(pagination.pageSizeOptions?.length ? pagination.pageSizeOptions : PAGE_SIZE_OPTIONS).map(
                      (n) => ({ key: String(n), text: String(n) })
                    )}
                    onChange={(_, opt) => setPageSize(Number(opt?.key) || 10)}
                    styles={{ root: { maxWidth: 120 } }}
                  />
                  <ChoiceGroup
                    label="Layout da paginação"
                    options={PAGINATION_LAYOUT_OPTIONS}
                    selectedKey={paginationLayout}
                    onChange={(_, opt) => opt && setPaginationLayout(opt.key as TPaginationLayout)}
                  />
                </>
              )}
            </Stack>

            <Separator />

            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                Modos de visualização
              </Text>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Ex.: Todas (sem filtro), Minhas (Author/Id eq [Me]), ou filtros customizados. O usuário alterna entre eles na lista.
              </Text>
              <Dropdown
                label="Modo padrão"
                options={viewModeDefaultOptions}
                selectedKey={activeViewModeId}
                onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => opt && setActiveViewModeId(String(opt.key))}
                styles={{ root: { maxWidth: 280 } }}
              />
              {viewModes.map((m) => (
                <div key={m.id} style={{ padding: 10, border: '1px solid #edebe9', borderRadius: 6, background: viewModeEditingId === m.id ? '#f3f9ff' : '#fff' }}>
                  {viewModeEditingId === m.id ? (
                    <Stack tokens={{ childrenGap: 10 }}>
                      <TextField label="Nome" value={viewModeEditLabel} onChange={(_: React.FormEvent, v?: string) => setViewModeEditLabel(v ?? '')} />
                      {viewModeEditFilters.map((f, i) => (
                        <Stack key={i} horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
                          <TextField placeholder="Campo (ex: Author/Id)" value={f.field} onChange={(_: React.FormEvent, v?: string) => updateViewModeFilter(i, { field: v ?? '' })} styles={{ root: { flex: 1 } }} />
                          <Dropdown options={VIEW_MODE_OPERATORS} selectedKey={f.operator} onChange={(_: React.FormEvent, opt?: IDropdownOption) => opt && updateViewModeFilter(i, { operator: opt.key as TFilterOperator })} styles={{ root: { width: 120 } }} />
                          <TextField placeholder="Valor ou [Me]" value={f.value} onChange={(_: React.FormEvent, v?: string) => updateViewModeFilter(i, { value: v ?? '' })} styles={{ root: { flex: 1 } }} />
                          <IconButton iconProps={{ iconName: 'Delete' }} title="Remover filtro" onClick={() => removeViewModeFilter(i)} />
                        </Stack>
                      ))}
                      <DefaultButton text="Adicionar filtro" onClick={addViewModeFilter} />
                      <Stack horizontal tokens={{ childrenGap: 8 }}>
                        <PrimaryButton text="Salvar" onClick={saveViewModeEdit} />
                        <DefaultButton text="Cancelar" onClick={() => setViewModeEditingId(null)} />
                      </Stack>
                    </Stack>
                  ) : (
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                      <Stack tokens={{ childrenGap: 2 }}>
                        <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>{m.label}</Text>
                        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{viewModeFilterSummary(m.filters)}</Text>
                      </Stack>
                      <Stack horizontal tokens={{ childrenGap: 4 }}>
                        <IconButton iconProps={{ iconName: 'Edit' }} title="Editar" onClick={() => startViewModeEdit(m)} />
                        <IconButton iconProps={{ iconName: 'Delete' }} title="Remover" onClick={() => removeViewMode(m.id)} disabled={m.id === 'all' || m.id === 'mine'} />
                      </Stack>
                    </Stack>
                  )}
                </div>
              ))}
              {viewModeEditingId !== null && !viewModes.some((m) => m.id === viewModeEditingId) && (
                <div style={{ padding: 10, border: '1px solid #c7e0f4', borderRadius: 6, background: '#f3f9ff' }}>
                  <Stack tokens={{ childrenGap: 10 }}>
                    <TextField label="Nome" value={viewModeEditLabel} onChange={(_: React.FormEvent, v?: string) => setViewModeEditLabel(v ?? '')} />
                    {viewModeEditFilters.map((f, i) => (
                      <Stack key={i} horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
                        <TextField placeholder="Campo (ex: Author/Id)" value={f.field} onChange={(_: React.FormEvent, v?: string) => updateViewModeFilter(i, { field: v ?? '' })} styles={{ root: { flex: 1 } }} />
                        <Dropdown options={VIEW_MODE_OPERATORS} selectedKey={f.operator} onChange={(_: React.FormEvent, opt?: IDropdownOption) => opt && updateViewModeFilter(i, { operator: opt.key as TFilterOperator })} styles={{ root: { width: 120 } }} />
                        <TextField placeholder="Valor ou [Me]" value={f.value} onChange={(_: React.FormEvent, v?: string) => updateViewModeFilter(i, { value: v ?? '' })} styles={{ root: { flex: 1 } }} />
                        <IconButton iconProps={{ iconName: 'Delete' }} title="Remover filtro" onClick={() => removeViewModeFilter(i)} />
                      </Stack>
                    ))}
                    <DefaultButton text="Adicionar filtro" onClick={addViewModeFilter} />
                    <Stack horizontal tokens={{ childrenGap: 8 }}>
                      <PrimaryButton text="Adicionar modo" onClick={saveViewModeEdit} />
                      <DefaultButton text="Cancelar" onClick={() => setViewModeEditingId(null)} />
                    </Stack>
                  </Stack>
                </div>
              )}
              {viewModeEditingId === null && <DefaultButton text="Adicionar modo de visualização" onClick={startViewModeAdd} />}
            </Stack>

            <Separator />

            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                Colunas
              </Text>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Marque as colunas que deseja exibir. Para lookups e usuários, escolha o campo de exibição.
              </Text>
            </Stack>
            {options.map((o) => (
              <Stack
                key={o.meta.InternalName}
                horizontal
                tokens={{ childrenGap: 12 }}
                verticalAlign="center"
                styles={{ root: { padding: '8px 0', borderBottom: '1px solid #f3f2f1' } }}
              >
                <Checkbox
                  checked={o.selected}
                  onChange={() => toggle(o.meta.InternalName)}
                  ariaLabel={o.meta.Title}
                />
                <Stack tokens={{ childrenGap: 4 }} styles={{ root: { flex: 1 } }}>
                  <TextField
                    label="Rótulo"
                    value={o.label}
                    onChange={(_, v) => setLabel(o.meta.InternalName, v ?? '')}
                    disabled={!o.selected}
                    styles={{ root: { maxWidth: 280 } }}
                  />
                  {EXPANDABLE.indexOf(o.meta.MappedType) !== -1 && o.selected && (
                    <Dropdown
                      label="Campo expandido (lookup/user)"
                      selectedKey={o.expandField || 'Title'}
                      options={getExpandFieldOptions(o.meta)}
                      onChange={(_, opt) => setExpandField(o.meta.InternalName, (opt?.key as string) ?? 'Title')}
                      styles={{ root: { maxWidth: 280 } }}
                    />
                  )}
                </Stack>
                <Text variant="small" styles={{ root: { color: '#a19f9d', minWidth: 80 } }}>
                  {o.meta.MappedType}
                </Text>
              </Stack>
            ))}
          </Stack>
        )}
      </div>
    </Panel>
  );
};
