import * as React from 'react';
import { useState } from 'react';
import {
  Stack,
  Text,
  Dropdown,
  IDropdownOption,
  TextField,
  PrimaryButton,
  DefaultButton,
  IconButton,
} from '@fluentui/react';
import type { IListViewModeConfig, IListViewFilterConfig, TFilterOperator } from '../../../core/config/types';
import { IWizardFormState } from '../types';

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

interface IStep5Props {
  form: IWizardFormState;
  onChange: (partial: Partial<IWizardFormState>) => void;
}

export const Step5ViewModes: React.FC<IStep5Props> = ({ form, onChange }) => {
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editLabel, setEditLabel] = useState('');
  const [editFilters, setEditFilters] = useState<IListViewFilterConfig[]>([]);

  const viewModes = form.viewModes ?? [];
  const activeViewModeId = form.activeViewModeId ?? 'all';

  const handleDefaultChange = (_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption): void => {
    if (opt) onChange({ activeViewModeId: String(opt.key) });
  };

  const startAdd = (): void => {
    const id = `custom_${Date.now()}`;
    setEditLabel('Novo modo');
    setEditFilters([]);
    setEditingId(id);
  };

  const startEdit = (m: IListViewModeConfig): void => {
    setEditingId(m.id);
    setEditLabel(m.label);
    setEditFilters(m.filters?.length ? m.filters.slice() : []);
  };

  const saveEdit = (): void => {
    if (editingId === null) return;
    const next = viewModes.slice();
    let idx = -1;
    for (let i = 0; i < next.length; i++) { if (next[i].id === editingId) { idx = i; break; } }
    const entry: IListViewModeConfig = { id: editingId, label: editLabel.trim() || editingId, filters: editFilters };
    if (idx >= 0) next[idx] = entry;
    else next.push(entry);
    onChange({ viewModes: next });
    setEditingId(null);
  };

  const cancelEdit = (): void => setEditingId(null);

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
        {viewModes.map((m) => (
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
                  {editFilters.map((f, i) => (
                    <Stack key={i} horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
                      <TextField
                        placeholder="Campo (ex: Author/Id, Status)"
                        value={f.field}
                        onChange={(_: React.FormEvent, v?: string) => updateFilter(i, { field: v ?? '' })}
                        styles={{ root: { flex: 1 } }}
                      />
                      <Dropdown
                        options={OPERATOR_OPTIONS}
                        selectedKey={f.operator}
                        onChange={(_: React.FormEvent, opt?: IDropdownOption) => opt && updateFilter(i, { operator: opt.key as TFilterOperator })}
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
        ))}
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
                {editFilters.map((f, i) => (
                  <Stack key={i} horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
                    <TextField
                      placeholder="Campo (ex: Author/Id, Status)"
                      value={f.field}
                      onChange={(_: React.FormEvent, v?: string) => updateFilter(i, { field: v ?? '' })}
                      styles={{ root: { flex: 1 } }}
                    />
                    <Dropdown
                      options={OPERATOR_OPTIONS}
                      selectedKey={f.operator}
                      onChange={(_: React.FormEvent, opt?: IDropdownOption) => opt && updateFilter(i, { operator: opt.key as TFilterOperator })}
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
