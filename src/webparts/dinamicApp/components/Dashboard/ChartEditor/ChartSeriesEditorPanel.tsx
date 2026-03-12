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
import { IChartSeriesConfig, IDashboardCardFilter, TAggregateType, TFilterOperator, TChartType, TDashboardType } from '../../../core/config/types';
import { FieldsService } from '../../../../../services';
import type { IFieldMetadata } from '../../../../../services';
import { ChartTypeCard } from '../ChartTypeCard';

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

const DASHBOARD_TYPE_OPTIONS: IChoiceGroupOption[] = [
  { key: 'cards', text: 'Cards' },
  { key: 'charts', text: 'Gráficos' },
];

const CHART_TYPES: TChartType[] = ['bar', 'line', 'area', 'pie', 'donut'];

interface ISeriesFormState {
  label: string;
  aggregate: TAggregateType;
  field: string;
  hasFilter: boolean;
  filters: IDashboardCardFilter[];
  color: string;
}

function initSeriesState(series?: IChartSeriesConfig): ISeriesFormState {
  const filters =
    series?.filters && series.filters.length > 0
      ? series.filters.slice()
      : series?.filter
        ? [{ field: series.filter.field, operator: series.filter.operator, value: series.filter.value }]
        : [{ field: '', operator: 'eq' as TFilterOperator, value: '' }];
  return {
    label: series?.label ?? '',
    aggregate: series?.aggregate ?? 'count',
    field: series?.field ?? '',
    hasFilter: true,
    filters,
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
  if (state.hasFilter && state.filters.length > 0) {
    const valid = state.filters.filter((f) => f.field.trim().length > 0 && String(f.value).trim().length > 0);
    if (valid.length > 0) s.filters = valid.map((f) => ({ field: f.field.trim(), operator: f.operator, value: String(f.value).trim() }));
  }
  return s;
}

type TPanelView = 'list' | 'form';

export interface IChartSeriesEditorSaveOptions {
  dashboardType?: TDashboardType;
  chartType?: TChartType;
}

interface IChartSeriesEditorPanelProps {
  isOpen: boolean;
  listTitle: string;
  series: IChartSeriesConfig[];
  dashboardType: TDashboardType;
  chartType?: TChartType;
  onSave: (series: IChartSeriesConfig[], options?: IChartSeriesEditorSaveOptions) => void;
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
          <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Filtros</Text>
          {state.filters.map((f, i) => (
            <Stack key={i} horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
              <Dropdown
                label="Campo"
                placeholder="Selecione"
                options={filterFieldOptions}
                selectedKey={f.field || ''}
                onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                  const next = state.filters.slice();
                  next[i] = { ...next[i], field: opt ? String(opt.key) : '' };
                  update({ filters: next });
                }}
                styles={{ root: { flex: 1 } }}
                disabled={fieldsLoading}
              />
              <Dropdown
                label="Operador"
                options={OPERATOR_OPTIONS}
                selectedKey={f.operator}
                onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                  if (opt) {
                    const next = state.filters.slice();
                    next[i] = { ...next[i], operator: String(opt.key) as TFilterOperator };
                    update({ filters: next });
                  }
                }}
                styles={{ root: { width: 140 } }}
              />
              <TextField
                label="Valor"
                value={f.value}
                onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) => {
                  const next = state.filters.slice();
                  next[i] = { ...next[i], value: v ?? '' };
                  update({ filters: next });
                }}
                placeholder="Ex: Ativo, [me]"
                styles={{ root: { flex: 1 } }}
              />
              <IconButton iconProps={{ iconName: 'Delete' }} title="Remover filtro" onClick={() => update({ filters: state.filters.filter((_, idx) => idx !== i) })} />
            </Stack>
          ))}
          <DefaultButton text="Adicionar filtro" onClick={() => update({ filters: [...state.filters, { field: '', operator: 'eq', value: '' }] })} />
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
  dashboardType,
  chartType = 'bar',
  onSave,
  onDismiss,
}) => {
  const [localSeries, setLocalSeries] = useState<IChartSeriesConfig[]>(() => [...series]);
  const [localDashboardType, setLocalDashboardType] = useState<TDashboardType>(dashboardType);
  const [localChartType, setLocalChartType] = useState<TChartType>(chartType);
  const [view, setView] = useState<TPanelView>('list');
  const [editingIndex, setEditingIndex] = useState<number | undefined>(undefined);

  useEffect(() => {
    if (isOpen) {
      setLocalSeries([...series]);
      setLocalDashboardType(dashboardType);
      setLocalChartType(chartType ?? 'bar');
      setView('list');
      setEditingIndex(undefined);
    }
  }, [isOpen, series, dashboardType, chartType]);

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
              <PrimaryButton
                text="Salvar"
                onClick={() => onSave(localSeries, { dashboardType: localDashboardType, chartType: localChartType })}
              />
              <DefaultButton text="Cancelar" onClick={onDismiss} />
            </>
          )}
        </Stack>
      )}
    >
      <div style={{ paddingTop: 16 }}>
        {view === 'list' && (
          <Stack tokens={{ childrenGap: 16 }}>
            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                Visualização do dashboard
              </Text>
              <ChoiceGroup
                options={DASHBOARD_TYPE_OPTIONS}
                selectedKey={localDashboardType}
                onChange={(_, opt) => opt && setLocalDashboardType(opt.key as TDashboardType)}
              />
            </Stack>
            {localDashboardType === 'charts' && (
              <Stack tokens={{ childrenGap: 10 }}>
                <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                  Escolha o tipo de gráfico
                </Text>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10 }}>
                  {CHART_TYPES.map((type) => (
                    <ChartTypeCard
                      key={type}
                      type={type}
                      selected={localChartType === type}
                      onClick={() => setLocalChartType(type)}
                    />
                  ))}
                </div>
              </Stack>
            )}
            <Separator />
            <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
              Séries
            </Text>
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
