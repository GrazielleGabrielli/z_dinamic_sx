import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import {
  Stack,
  Text,
  Dropdown,
  IDropdownOption,
  TextField,
  PrimaryButton,
  DefaultButton,
  IconButton,
  Spinner,
  SpinnerSize,
  ChoiceGroup,
  IChoiceGroupOption,
} from '@fluentui/react';
import { FieldsService } from '../../../../../services';
import type { IFieldMetadata } from '../../../../../services';
import type {
  IListViewModeConfig,
  IListViewFilterConfig,
  IListViewModeAccessConfig,
  TFilterOperator,
  TViewModePicker,
} from '../../../core/config/types';
import { ViewModeAccessSection, accessSummary } from '../../shared/ViewModeAccessSection';
import { isNoteFieldMeta } from '../../../core/listView';
import { IWizardFormState } from '../types';

const EXPANDABLE = ['lookup', 'lookupmulti', 'user', 'usermulti'];
const SIMPLE_FIELD_TYPES = ['text', 'multiline', 'number', 'currency', 'boolean', 'choice', 'multichoice', 'datetime', 'url'];
const USER_EXPAND_FIELDS: IDropdownOption[] = [
  { key: 'Id', text: 'Id' },
  { key: 'Title', text: 'Title' },
  { key: 'EMail', text: 'EMail' },
  { key: 'LoginName', text: 'LoginName' },
];

function buildExpandOptionsFromLookupList(fields: IFieldMetadata[]): IDropdownOption[] {
  const simple = fields.filter(
    (f) =>
      SIMPLE_FIELD_TYPES.indexOf(f.MappedType) !== -1 &&
      f.InternalName !== 'Id' &&
      f.InternalName !== 'Title' &&
      !isNoteFieldMeta(f)
  );
  const options: IDropdownOption[] = [
    { key: 'Id', text: 'Id' },
    { key: 'Title', text: 'Title' },
  ];
  simple.forEach((f) => options.push({ key: f.InternalName, text: `${f.Title} (${f.InternalName})` }));
  return options;
}

const OPERATOR_OPTIONS: IDropdownOption[] = [
  { key: 'eq', text: 'Igual a' },
  { key: 'ne', text: 'Diferente de' },
  { key: 'contains', text: 'Contém' },
  { key: 'gt', text: 'Maior que' },
  { key: 'ge', text: 'Maior ou igual' },
  { key: 'lt', text: 'Menor que' },
  { key: 'le', text: 'Menor ou igual' },
];

function filterSummary(filters: IListViewFilterConfig[]): string {
  if (!filters || filters.length === 0) return 'Sem filtros';
  return filters.map((f) => `${f.field} ${f.operator} "${f.value}"`).join(' e ');
}

const VIEW_MODE_PICKER_OPTIONS: IChoiceGroupOption[] = [
  { key: 'dropdown', text: 'Lista suspensa' },
  { key: 'tabs', text: 'Abas horizontais' },
];

interface IStep5Props {
  form: IWizardFormState;
  listTitle: string;
  listWebServerRelativeUrl?: string;
  pageWebServerRelativeUrl: string;
  onChange: (partial: Partial<IWizardFormState>) => void;
}

export const Step5ViewModes: React.FC<IStep5Props> = ({
  form,
  listTitle,
  listWebServerRelativeUrl,
  pageWebServerRelativeUrl,
  onChange,
}) => {
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editLabel, setEditLabel] = useState('');
  const [editFilters, setEditFilters] = useState<IListViewFilterConfig[]>([]);
  const [editAccess, setEditAccess] = useState<IListViewModeAccessConfig | undefined>(undefined);
  const [listFields, setListFields] = useState<IFieldMetadata[]>([]);
  const [lookupListFields, setLookupListFields] = useState<Record<string, IFieldMetadata[]>>({});
  const [fieldsLoading, setFieldsLoading] = useState(false);

  useEffect(() => {
    if (!listTitle || !listTitle.trim()) {
      setListFields([]);
      setLookupListFields({});
      return;
    }
    setFieldsLoading(true);
    const svc = new FieldsService();
    const lw = listWebServerRelativeUrl?.trim() || undefined;
    svc
      .getVisibleFields(listTitle.trim(), lw)
      .then((fields) => {
        setListFields(fields);
        const listIds = fields
          .filter((f) => EXPANDABLE.indexOf(f.MappedType) !== -1 && f.LookupList)
          .map((f) => f.LookupList as string);
        const uniqueIds = listIds.filter((id, i) => listIds.indexOf(id) === i);
        return Promise.all(
          uniqueIds.map((id) => svc.getFields(id, lw).then((listFields) => ({ id, listFields })))
        );
      })
      .then((results) => {
        const next: Record<string, IFieldMetadata[]> = {};
        results.forEach((r: { id: string; listFields: IFieldMetadata[] }) => { next[r.id] = r.listFields; });
        setLookupListFields(next);
      })
      .then(() => setFieldsLoading(false), () => setFieldsLoading(false));
  }, [listTitle, listWebServerRelativeUrl]);

  const filterFieldOptions = useMemo((): IDropdownOption[] => {
    const empty: IDropdownOption = { key: '', text: '— selecione —' };
    const rest: IDropdownOption[] = [];
    for (let i = 0; i < listFields.length; i++) {
      const f = listFields[i];
      if (isNoteFieldMeta(f)) continue;
      if (EXPANDABLE.indexOf(f.MappedType) === -1) {
        rest.push({ key: f.InternalName, text: `${f.Title} (${f.InternalName})` });
      } else {
        const expandOpts =
          f.MappedType === 'user' || f.MappedType === 'usermulti'
            ? USER_EXPAND_FIELDS
            : f.LookupList && lookupListFields[f.LookupList]
              ? buildExpandOptionsFromLookupList(lookupListFields[f.LookupList])
              : [{ key: 'Title', text: 'Title' }, { key: 'Id', text: 'Id' }];
        for (let j = 0; j < expandOpts.length; j++) {
          const opt = expandOpts[j];
          rest.push({
            key: `${f.InternalName}/${String(opt.key)}`,
            text: `${f.Title} – ${opt.text}`,
          });
        }
      }
    }
    return [empty, ...rest];
  }, [listFields, lookupListFields]);

  const viewModes = form.viewModes ?? [];
  const activeViewModeId = form.activeViewModeId ?? 'all';

  const handleDefaultChange = (_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption): void => {
    if (opt) onChange({ activeViewModeId: String(opt.key) });
  };

  const startAdd = (): void => {
    const id = `custom_${Date.now()}`;
    setEditLabel('Novo modo');
    setEditFilters([]);
    setEditAccess(undefined);
    setEditingId(id);
  };

  const startEdit = (m: IListViewModeConfig): void => {
    setEditingId(m.id);
    setEditLabel(m.label);
    setEditFilters(m.filters?.length ? m.filters.slice() : []);
    setEditAccess(m.access);
  };

  const saveEdit = (): void => {
    if (editingId === null) return;
    const next = viewModes.slice();
    let idx = -1;
    for (let i = 0; i < next.length; i++) { if (next[i].id === editingId) { idx = i; break; } }
    const entry: IListViewModeConfig = {
      id: editingId,
      label: editLabel.trim() || editingId,
      filters: editFilters,
      ...(editAccess !== undefined ? { access: editAccess } : {}),
    };
    if (idx >= 0) next[idx] = entry;
    else next.push(entry);
    onChange({ viewModes: next });
    setEditingId(null);
  };

  const cancelEdit = (): void => {
    setEditingId(null);
    setEditAccess(undefined);
  };

  const removeMode = (id: string): void => {
    if (id === 'all' || id === 'mine') return;
    const next = viewModes.filter((m) => m.id !== id);
    const nextActive = activeViewModeId === id ? (next[0]?.id ?? 'all') : activeViewModeId;
    onChange({ viewModes: next, activeViewModeId: nextActive });
  };

  const addFilter = (): void => setEditFilters([...editFilters, { field: '', operator: 'eq', value: '' }]);
  const removeFilter = (i: number): void => setEditFilters(editFilters.filter((_, idx) => idx !== i));
  const updateFilter = (i: number, part: Partial<IListViewFilterConfig>): void => {
    const next = editFilters.slice();
    next[i] = { ...next[i], ...part };
    setEditFilters(next);
  };

  const defaultOptions: IDropdownOption[] = viewModes.map((m) => ({ key: m.id, text: m.label }));

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <Stack.Item>
        <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
          Modos de visualização
        </Text>
        <Text variant="medium" styles={{ root: { color: '#605e5c', marginTop: 4, display: 'block' } }}>
          Defina opções como &quot;Todas&quot;, &quot;Minhas&quot; (itens do usuário atual) e outros filtros. O usuário poderá alternar entre eles na lista.
        </Text>
      </Stack.Item>

      <ChoiceGroup
        label="Controlo na lista"
        selectedKey={form.viewModePicker}
        options={VIEW_MODE_PICKER_OPTIONS}
        onChange={(_, opt) => {
          const k = (opt?.key as string | undefined) ?? 'dropdown';
          const next: TViewModePicker = k === 'tabs' ? 'tabs' : 'dropdown';
          onChange({ viewModePicker: next });
        }}
        styles={{
          flexContainer: { display: 'flex', flexWrap: 'wrap', columnGap: '12px', rowGap: '4px' },
        }}
      />

      <Dropdown
        label="Modo de visualização padrão"
        options={defaultOptions}
        selectedKey={activeViewModeId}
        onChange={handleDefaultChange}
        styles={{ root: { maxWidth: 320 } }}
      />

      <Stack tokens={{ childrenGap: 8 }}>
        <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
          Modos disponíveis
        </Text>
        {viewModes.map((m) => {
          const accessLine = accessSummary(m.access);
          return (
          <div
            key={m.id}
            style={{
              padding: 12,
              border: '1px solid #edebe9',
              borderRadius: 8,
              background: editingId === m.id ? '#f3f9ff' : '#fff',
            }}
          >
            {editingId === m.id ? (
              <Stack tokens={{ childrenGap: 12 }}>
                <TextField
                  label="Nome do modo"
                  value={editLabel}
                  onChange={(_: React.FormEvent, v?: string) => setEditLabel(v ?? '')}
                />
                <Stack tokens={{ childrenGap: 8 }}>
                  <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                    Filtros
                  </Text>
                  {fieldsLoading && (
                    <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                      <Spinner size={SpinnerSize.small} />
                      <Text variant="small">Carregando campos...</Text>
                    </Stack>
                  )}
                  {editFilters.map((f, i) => (
                    <Stack key={i} horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
                      <Dropdown
                        placeholder="Campo"
                        options={filterFieldOptions}
                        selectedKey={f.field || ''}
                        onChange={(_: React.FormEvent, opt?: IDropdownOption) => updateFilter(i, { field: (opt?.key as string) ?? '' })}
                        styles={{ root: { flex: 1 } }}
                        disabled={fieldsLoading}
                      />
                      <Dropdown
                        options={OPERATOR_OPTIONS}
                        selectedKey={f.operator}
                        onChange={(_: React.FormEvent, opt?: IDropdownOption) => opt != null && updateFilter(i, { operator: String(opt.key) as TFilterOperator })}
                        styles={{ root: { width: 140 } }}
                      />
                      <TextField
                        placeholder="Valor ou [Me]"
                        value={f.value}
                        onChange={(_: React.FormEvent, v?: string) => updateFilter(i, { value: v ?? '' })}
                        styles={{ root: { flex: 1 } }}
                      />
                      <IconButton iconProps={{ iconName: 'Delete' }} title="Remover filtro" onClick={() => removeFilter(i)} />
                    </Stack>
                  ))}
                  <DefaultButton text="Adicionar filtro" onClick={addFilter} />
                </Stack>
                <ViewModeAccessSection
                  value={editAccess}
                  onChange={setEditAccess}
                  pageWebServerRelativeUrl={pageWebServerRelativeUrl}
                  listWebServerRelativeUrl={listWebServerRelativeUrl}
                />
                <Stack horizontal tokens={{ childrenGap: 8 }}>
                  <PrimaryButton text="Salvar" onClick={saveEdit} />
                  <DefaultButton text="Cancelar" onClick={cancelEdit} />
                </Stack>
              </Stack>
            ) : (
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Stack tokens={{ childrenGap: 2 }}>
                  <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                    {m.label}
                  </Text>
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    {filterSummary(m.filters)}
                  </Text>
                  {accessLine ? (
                    <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                      {accessLine}
                    </Text>
                  ) : null}
                </Stack>
                <Stack horizontal tokens={{ childrenGap: 4 }}>
                  <IconButton iconProps={{ iconName: 'Edit' }} title="Editar" onClick={() => startEdit(m)} />
                  <IconButton
                    iconProps={{ iconName: 'Delete' }}
                    title="Remover"
                    onClick={() => removeMode(m.id)}
                    disabled={m.id === 'all' || m.id === 'mine'}
                  />
                </Stack>
              </Stack>
            )}
          </div>
        );
        })}
        {editingId !== null && !viewModes.some((m) => m.id === editingId) && (
          <div style={{ padding: 12, border: '1px solid #c7e0f4', borderRadius: 8, background: '#f3f9ff' }}>
            <Stack tokens={{ childrenGap: 12 }}>
              <TextField
                label="Nome do modo"
                value={editLabel}
                onChange={(_: React.FormEvent, v?: string) => setEditLabel(v ?? '')}
              />
              <Stack tokens={{ childrenGap: 8 }}>
                <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                  Filtros
                </Text>
                {fieldsLoading && (
                  <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                    <Spinner size={SpinnerSize.small} />
                    <Text variant="small">Carregando campos...</Text>
                  </Stack>
                )}
                {editFilters.map((f, i) => (
                  <Stack key={i} horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
                    <Dropdown
                      placeholder="Campo"
                      options={filterFieldOptions}
                      selectedKey={f.field || ''}
                      onChange={(_: React.FormEvent, opt?: IDropdownOption) => updateFilter(i, { field: (opt?.key as string) ?? '' })}
                      styles={{ root: { flex: 1 } }}
                      disabled={fieldsLoading}
                    />
                    <Dropdown
                      options={OPERATOR_OPTIONS}
                      selectedKey={f.operator}
                      onChange={(_: React.FormEvent, opt?: IDropdownOption) => opt != null && updateFilter(i, { operator: String(opt.key) as TFilterOperator })}
                      styles={{ root: { width: 140 } }}
                    />
                    <TextField
                      placeholder="Valor ou [Me]"
                      value={f.value}
                      onChange={(_: React.FormEvent, v?: string) => updateFilter(i, { value: v ?? '' })}
                      styles={{ root: { flex: 1 } }}
                    />
                    <IconButton iconProps={{ iconName: 'Delete' }} title="Remover filtro" onClick={() => removeFilter(i)} />
                  </Stack>
                ))}
                <DefaultButton text="Adicionar filtro" onClick={addFilter} />
              </Stack>
              <ViewModeAccessSection
                value={editAccess}
                onChange={setEditAccess}
                pageWebServerRelativeUrl={pageWebServerRelativeUrl}
                listWebServerRelativeUrl={listWebServerRelativeUrl}
              />
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <PrimaryButton text="Adicionar modo" onClick={saveEdit} />
                <DefaultButton text="Cancelar" onClick={cancelEdit} />
              </Stack>
            </Stack>
          </div>
        )}
        {editingId === null && (
          <DefaultButton text="Adicionar modo de visualização" onClick={startAdd} />
        )}
      </Stack>
    </Stack>
  );
};
