import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  TextField,
  Toggle,
  Dropdown,
  IDropdownOption,
  ChoiceGroup,
  IChoiceGroupOption,
  PrimaryButton,
  DefaultButton,
  IconButton,
  Separator,
  Spinner,
  SpinnerSize,
} from '@fluentui/react';
import { IChartSeriesConfig, TAggregateType, TFilterOperator } from '../../../core/config/types';
import { FieldsService } from '../../../../../services';
import type { IFieldMetadata } from '../../../../../services';

const NUMERIC_MAPPED_TYPES: string[] = ['number', 'currency', 'calculated'];

function isNumericField(f: IFieldMetadata): boolean {
  return NUMERIC_MAPPED_TYPES.indexOf(f.MappedType) !== -1;
}

const AGGREGATE_OPTIONS: IChoiceGroupOption[] = [
  { key: 'count', text: 'Contagem — conta os itens' },
  { key: 'sum', text: 'Soma — soma um campo numérico' },
];

const OPERATOR_OPTIONS: IDropdownOption[] = [
  { key: 'eq', text: 'igual a' },
  { key: 'ne', text: 'diferente de' },
  { key: 'gt', text: 'maior que' },
  { key: 'lt', text: 'menor que' },
  { key: 'ge', text: 'maior ou igual a' },
  { key: 'le', text: 'menor ou igual a' },
  { key: 'contains', text: 'contém (texto)' },
];

const PRESET_COLORS = [
  '#0078d4', '#2b88d8', '#71afe5',
  '#00b294', '#ffaa44', '#d13438',
  '#8764b8', '#038387', '#ca5010',
];

interface ISeriesFormState {
  label: string;
  aggregate: TAggregateType;
  field: string;
  hasFilter: boolean;
  filterField: string;
  filterOperator: TFilterOperator;
  filterValue: string;
  color: string;
}

function initSeriesState(series?: IChartSeriesConfig): ISeriesFormState {
  return {
    label: series?.label ?? '',
    aggregate: series?.aggregate ?? 'count',
    field: series?.field ?? '',
    hasFilter: series?.filter !== undefined,
    filterField: series?.filter?.field ?? '',
    filterOperator: series?.filter?.operator ?? 'eq',
    filterValue: series?.filter?.value ?? '',
    color: series?.color ?? PRESET_COLORS[0],
  };
}

function buildSeries(state: ISeriesFormState, existingId?: string): IChartSeriesConfig {
  const s: IChartSeriesConfig = {
    id: existingId ?? `series_${String(Date.now())}`,
    label: state.label.trim(),
    aggregate: state.aggregate,
    color: state.color || undefined,
  };
  if (state.aggregate === 'sum' && state.field.trim()) {
    s.field = state.field.trim();
  }
  if (state.hasFilter && state.filterField.trim() && state.filterValue.trim()) {
    s.filter = {
      field: state.filterField.trim(),
      operator: state.filterOperator,
      value: state.filterValue.trim(),
    };
  }
  return s;
}

type TPanelView = 'list' | 'form';

interface IChartSeriesEditorPanelProps {
  isOpen: boolean;
  listTitle: string;
  series: IChartSeriesConfig[];
  onSave: (series: IChartSeriesConfig[]) => void;
  onDismiss: () => void;
}

const SeriesForm: React.FC<{
  listTitle: string;
  series?: IChartSeriesConfig;
  onConfirm: (s: IChartSeriesConfig) => void;
  onBack: () => void;
}> = ({ listTitle, series, onConfirm, onBack }) => {
  const [state, setState] = useState<ISeriesFormState>(() => initSeriesState(series));
  const [listFields, setListFields] = useState<IFieldMetadata[]>([]);
  const [fieldsLoading, setFieldsLoading] = useState(false);
  const [fieldsError, setFieldsError] = useState<string | undefined>(undefined);

  const update = (partial: Partial<ISeriesFormState>): void => {
    setState((prev) => ({ ...prev, ...partial }));
  };

  useEffect(() => {
    if (!listTitle.trim()) return;
    setFieldsLoading(true);
    setFieldsError(undefined);
    const svc = new FieldsService();
    svc.getVisibleFields(listTitle.trim())
      .then((fields) => { setListFields(fields); setFieldsLoading(false); })
      .catch((err) => {
        setFieldsError(err instanceof Error ? err.message : String(err));
        setFieldsLoading(false);
      });
  }, [listTitle]);

  const numericFields = useMemo(() => listFields.filter(isNumericField), [listFields]);

  const filterFieldOptions = useMemo((): IDropdownOption[] => ([
    { key: '', text: '— selecione —' },
    ...listFields.map((f) => ({ key: f.InternalName, text: `${f.Title} (${f.InternalName})` })),
  ]), [listFields]);

  const sumFieldOptions = useMemo((): IDropdownOption[] => ([
    { key: '', text: '— selecione —' },
    ...numericFields.map((f) => ({ key: f.InternalName, text: `${f.Title} (${f.InternalName})` })),
  ]), [numericFields]);

  const isValid =
    state.label.trim().length > 0 &&
    (state.aggregate === 'count' || state.field.trim().length > 0);

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <TextField
        label="Rótulo da série"
        value={state.label}
        onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) => update({ label: v ?? '' })}
        required
        placeholder="Ex: Itens Pendentes, Concluídos..."
      />

      <ChoiceGroup
        label="Tipo de agregação"
        options={AGGREGATE_OPTIONS}
        selectedKey={state.aggregate}
        onChange={(_: React.FormEvent<HTMLElement | HTMLInputElement> | undefined, opt?: IChoiceGroupOption) => {
          if (opt) update({ aggregate: opt.key as TAggregateType, field: '' });
        }}
      />

      {state.aggregate === 'sum' && (
        <Dropdown
          label="Campo numérico"
          placeholder="Selecione o campo"
          options={sumFieldOptions}
          selectedKey={state.field || ''}
          onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) =>
            update({ field: opt ? String(opt.key) : '' })
          }
          required
          disabled={fieldsLoading}
        />
      )}

      <Separator />

      <Toggle
        label="Aplicar filtro nos dados"
        checked={state.hasFilter}
        onChange={(_: React.MouseEvent<HTMLElement>, checked?: boolean) => update({ hasFilter: !!checked })}
        onText="Sim"
        offText="Não"
      />

      {state.hasFilter && (
        <Stack
          tokens={{ childrenGap: 12 }}
          styles={{ root: { background: '#faf9f8', padding: 16, borderRadius: 6, border: '1px solid #edebe9' } }}
        >
          {fieldsLoading && (
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
              <Spinner size={SpinnerSize.small} />
              <Text variant="small">Carregando campos...</Text>
            </Stack>
          )}
          {fieldsError && (
            <Text variant="small" styles={{ root: { color: '#d13438' } }}>{fieldsError}</Text>
          )}
          <Dropdown
            label="Campo do filtro"
            placeholder="Selecione o campo"
            options={filterFieldOptions}
            selectedKey={state.filterField || ''}
            onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) =>
              update({ filterField: opt ? String(opt.key) : '' })
            }
            disabled={fieldsLoading}
          />
          <Dropdown
            label="Operador"
            options={OPERATOR_OPTIONS}
            selectedKey={state.filterOperator}
            onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
              if (opt) update({ filterOperator: opt.key as TFilterOperator });
            }}
          />
          <TextField
            label="Valor"
            value={state.filterValue}
            onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) =>
              update({ filterValue: v ?? '' })
            }
            placeholder="Ex: Ativo, Pendente, 100"
          />
        </Stack>
      )}

      <Separator />

      <Stack tokens={{ childrenGap: 8 }}>
        <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Cor da série</Text>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          {PRESET_COLORS.map((c) => (
            <div
              key={c}
              onClick={() => update({ color: c })}
              style={{
                width: 28,
                height: 28,
                borderRadius: '50%',
                background: c,
                cursor: 'pointer',
                border: state.color === c ? '3px solid #323130' : '3px solid transparent',
                boxSizing: 'border-box',
              }}
            />
          ))}
        </div>
        <TextField
          label="Cor personalizada (hex)"
          value={state.color}
          onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) =>
            update({ color: v ?? '' })
          }
          placeholder="#0078d4"
          styles={{ root: { maxWidth: 180 } }}
        />
      </Stack>

      <Separator />

      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton text="Confirmar" onClick={() => isValid && onConfirm(buildSeries(state, series?.id))} disabled={!isValid} />
        <DefaultButton text="Voltar" onClick={onBack} />
      </Stack>
    </Stack>
  );
};

export const ChartSeriesEditorPanel: React.FC<IChartSeriesEditorPanelProps> = ({
  isOpen,
  listTitle,
  series,
  onSave,
  onDismiss,
}) => {
  const [localSeries, setLocalSeries] = useState<IChartSeriesConfig[]>(() => [...series]);
  const [view, setView] = useState<TPanelView>('list');
  const [editingIndex, setEditingIndex] = useState<number | undefined>(undefined);

  useEffect(() => {
    if (isOpen) {
      setLocalSeries([...series]);
      setView('list');
      setEditingIndex(undefined);
    }
  }, [isOpen]);

  const handleEdit = (index: number): void => {
    setEditingIndex(index);
    setView('form');
  };

  const handleDelete = (index: number): void => {
    setLocalSeries((prev) => prev.filter((_, i) => i !== index));
  };

  const handleConfirm = (s: IChartSeriesConfig): void => {
    if (editingIndex !== undefined) {
      setLocalSeries((prev) => prev.map((item, i) => (i === editingIndex ? s : item)));
    } else {
      setLocalSeries((prev) => [...prev, s]);
    }
    setView('list');
  };

  const panelHeader = view === 'list'
    ? 'Séries do gráfico'
    : editingIndex !== undefined ? 'Editar série' : 'Nova série';

  const editingSeries = editingIndex !== undefined ? localSeries[editingIndex] : undefined;

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText={panelHeader}
      closeButtonAriaLabel="Fechar"
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { paddingBottom: 16 } }}>
          {view === 'list' && (
            <>
              <PrimaryButton text="Salvar" onClick={() => onSave(localSeries)} />
              <DefaultButton text="Cancelar" onClick={onDismiss} />
            </>
          )}
        </Stack>
      )}
    >
      <div style={{ paddingTop: 16 }}>
        {view === 'list' && (
          <Stack tokens={{ childrenGap: 0 }}>
            {localSeries.length === 0 && (
              <Text variant="small" styles={{ root: { color: '#a19f9d', padding: '16px 0' } }}>
                Nenhuma série configurada ainda.
              </Text>
            )}
            {localSeries.map((s, index) => (
              <React.Fragment key={s.id}>
                <div style={{ padding: '14px 0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
                    <div style={{ width: 14, height: 14, borderRadius: '50%', background: s.color ?? '#0078d4', flexShrink: 0 }} />
                    <Stack tokens={{ childrenGap: 2 }}>
                      <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>{s.label}</Text>
                      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                        {s.aggregate === 'count' ? 'contagem' : `soma · ${s.field ?? ''}`}
                        {s.filter !== undefined ? ` · filtro: ${s.filter.field} ${s.filter.operator} "${s.filter.value}"` : ''}
                      </Text>
                    </Stack>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 2 }}>
                    <IconButton iconProps={{ iconName: 'Edit' }} title="Editar" onClick={() => handleEdit(index)} styles={{ root: { color: '#0078d4' } }} />
                    <IconButton iconProps={{ iconName: 'Delete' }} title="Remover" onClick={() => handleDelete(index)} styles={{ root: { color: '#d13438' } }} />
                  </Stack>
                </div>
                {index < localSeries.length - 1 && <Separator styles={{ root: { padding: 0 } }} />}
              </React.Fragment>
            ))}
            <div style={{ marginTop: 20 }}>
              <DefaultButton
                iconProps={{ iconName: 'Add' }}
                text="Adicionar série"
                onClick={() => { setEditingIndex(undefined); setView('form'); }}
              />
            </div>
          </Stack>
        )}

        {view === 'form' && (
          <SeriesForm
            listTitle={listTitle}
            series={editingSeries}
            onConfirm={handleConfirm}
            onBack={() => setView('list')}
          />
        )}
      </div>
    </Panel>
  );
};
