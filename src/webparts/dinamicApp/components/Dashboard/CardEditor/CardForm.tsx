import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import {
  Stack,
  Text,
  TextField,
  ChoiceGroup,
  IChoiceGroupOption,
  Toggle,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  Separator,
  Pivot,
  PivotItem,
  Spinner,
  SpinnerSize,
  IconButton,
} from '@fluentui/react';
import {
  IDashboardCardConfig,
  IDashboardCardFilter,
  IDashboardCardLayoutStyle,
  IDashboardCardStyleConfig,
  TAggregateType,
  TFilterOperator,
} from '../../../core/config/types';
import { mergeWithDefaultStyle, resolveEffectiveCardStyle } from '../../../core/dashboard/utils/dashboardCardStyles';
import { IDashboardCardResult } from '../../../core/dashboard/types';
import { DashboardCard } from '../DashboardCard';
import { FieldsService } from '../../../../../services';
import type { IFieldMetadata } from '../../../../../services';

const NUMERIC_MAPPED_TYPES: string[] = ['number', 'currency', 'calculated'];

function isNumericField(f: IFieldMetadata): boolean {
  return NUMERIC_MAPPED_TYPES.indexOf(f.MappedType) !== -1;
}

interface ICardFormProps {
  listTitle: string;
  listWebServerRelativeUrl?: string;
  card: IDashboardCardConfig | undefined;
  cardLayoutStyle: IDashboardCardLayoutStyle;
  onConfirm: (card: IDashboardCardConfig) => void;
  onBack: () => void;
}

interface ICardFormState {
  title: string;
  subtitle: string;
  emptyValueText: string;
  errorText: string;
  loadingText: string;
  aggregate: TAggregateType;
  field: string;
  hasFilter: boolean;
  filters: IDashboardCardFilter[];
  backgroundColor: string;
  borderColor: string;
  titleColor: string;
  subtitleColor: string;
  valueColor: string;
  iconName: string;
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

const ICON_OPTIONS: IDropdownOption[] = [
  { key: '', text: '— nenhum —' },
  { key: 'NumberField', text: 'Número' },
  { key: 'Money', text: 'Dinheiro' },
  { key: 'People', text: 'Pessoas' },
  { key: 'Tag', text: 'Etiqueta' },
  { key: 'CheckMark', text: 'Concluído' },
  { key: 'Warning', text: 'Alerta' },
  { key: 'Clock', text: 'Tempo' },
  { key: 'Database', text: 'Dados' },
  { key: 'Info', text: 'Informação' },
  { key: 'Filter', text: 'Filtro' },
  { key: 'Chart', text: 'Gráfico' },
  { key: 'Add', text: 'Adicionar' },
  { key: 'StatusCircleCheckmark', text: 'Aprovado' },
  { key: 'Cancel', text: 'Cancelado' },
];

function initState(card: IDashboardCardConfig | undefined, layout: IDashboardCardLayoutStyle): ICardFormState {
  const s = mergeWithDefaultStyle(card?.style, layout);
  return {
    title: card?.title ?? '',
    subtitle: card?.subtitle ?? '',
    emptyValueText: card?.emptyValueText ?? 'Nenhum item',
    errorText: card?.errorText ?? 'Erro ao carregar',
    loadingText: card?.loadingText ?? 'Carregando...',
    aggregate: card?.aggregate ?? 'count',
    field: card?.field ?? '',
    hasFilter: true,
    filters: (card?.filters && card.filters.length > 0)
      ? card.filters.slice()
      : card?.filter
        ? [{ field: card.filter.field, operator: card.filter.operator, value: card.filter.value }]
        : [{ field: '', operator: 'eq' as TFilterOperator, value: '' }],
    backgroundColor: s.backgroundColor ?? '',
    borderColor: s.borderColor ?? '',
    titleColor: s.titleColor ?? '',
    subtitleColor: s.subtitleColor ?? '',
    valueColor: s.valueColor ?? '',
    iconName: s.iconName ?? '',
  };
}

function buildCardStyle(
  state: ICardFormState,
  layout: IDashboardCardLayoutStyle
): Partial<IDashboardCardStyleConfig> {
  const style: Partial<IDashboardCardStyleConfig> = {};
  if (state.backgroundColor) style.backgroundColor = state.backgroundColor;
  if (state.borderColor) style.borderColor = state.borderColor;
  if (state.titleColor) style.titleColor = state.titleColor;
  if (state.subtitleColor) style.subtitleColor = state.subtitleColor;
  if (state.valueColor) style.valueColor = state.valueColor;
  if (layout.showIcon && state.iconName) style.iconName = state.iconName;
  return style;
}

function buildCard(
  state: ICardFormState,
  layout: IDashboardCardLayoutStyle,
  existingId?: string
): IDashboardCardConfig {
  const appearance = buildCardStyle(state, layout);
  const card: IDashboardCardConfig = {
    id: existingId ?? `card_${String(Date.now())}`,
    title: state.title.trim(),
    aggregate: state.aggregate,
    ...(Object.keys(appearance).length > 0 ? { style: appearance as IDashboardCardStyleConfig } : {}),
  };
  if (state.subtitle.trim()) card.subtitle = state.subtitle.trim();
  if (state.emptyValueText.trim()) card.emptyValueText = state.emptyValueText.trim();
  if (state.errorText.trim()) card.errorText = state.errorText.trim();
  if (state.loadingText.trim()) card.loadingText = state.loadingText.trim();
  if (state.aggregate === 'sum' && state.field.trim().length > 0) {
    card.field = state.field.trim();
  }
  if (state.hasFilter && state.filters.length > 0) {
    const valid = state.filters.filter((f) => f.field.trim().length > 0 && String(f.value).trim().length > 0);
    if (valid.length > 0) card.filters = valid.map((f) => ({ field: f.field.trim(), operator: f.operator, value: String(f.value).trim() }));
  }
  return card;
}

export const CardForm: React.FC<ICardFormProps> = ({
  listTitle,
  listWebServerRelativeUrl,
  card,
  cardLayoutStyle,
  onConfirm,
  onBack,
}) => {
  const lw = listWebServerRelativeUrl?.trim() || undefined;
  const [state, setState] = useState<ICardFormState>(() => initState(card, cardLayoutStyle));

  useEffect(() => {
    setState(initState(card, cardLayoutStyle));
  }, [card, cardLayoutStyle]);
  const [listFields, setListFields] = useState<IFieldMetadata[]>([]);
  const [fieldsLoading, setFieldsLoading] = useState(false);
  const [fieldsError, setFieldsError] = useState<string | undefined>(undefined);

  useEffect(() => {
    if (!listTitle || !listTitle.trim()) {
      setListFields([]);
      setFieldsError(undefined);
      return;
    }
    setFieldsLoading(true);
    setFieldsError(undefined);
    const svc = new FieldsService();
    svc
      .getVisibleFields(listTitle.trim(), lw)
      .then((fields) => {
        setListFields(fields);
        setFieldsLoading(false);
      })
      .catch((err) => {
        setListFields([]);
        setFieldsError(err instanceof Error ? err.message : String(err));
        setFieldsLoading(false);
      });
  }, [listTitle, lw]);

  const numericFields = useMemo(
    () => listFields.filter(isNumericField),
    [listFields]
  );
  const filterFieldOptions = useMemo((): IDropdownOption[] => {
    const source = state.aggregate === 'sum' ? numericFields : listFields;
    return [
      { key: '', text: '— selecione —' },
      ...source.map((f) => ({ key: f.InternalName, text: `${f.Title} (${f.InternalName})` })),
    ];
  }, [listFields, numericFields, state.aggregate]);
  const sumFieldOptions = useMemo((): IDropdownOption[] => {
    return [
      { key: '', text: '— selecione —' },
      ...numericFields.map((f) => ({ key: f.InternalName, text: `${f.Title} (${f.InternalName})` })),
    ];
  }, [numericFields]);

  const update = (partial: Partial<ICardFormState>): void => {
    setState((prev) => ({ ...prev, ...partial }));
  };

  const isValid =
    state.title.trim().length > 0 &&
    (state.aggregate === 'count' || state.field.trim().length > 0) &&
    (!cardLayoutStyle.showIcon || state.iconName.trim().length > 0);

  const handleConfirm = (): void => {
    if (!isValid) return;
    onConfirm(buildCard(state, cardLayoutStyle, card?.id));
  };

  const previewAppearance = buildCardStyle(state, cardLayoutStyle);
  const previewCardConfig: IDashboardCardConfig = {
    id: 'preview',
    title: state.title.trim() || 'Título do card',
    subtitle: state.subtitle.trim() || undefined,
    aggregate: state.aggregate,
    emptyValueText: state.emptyValueText,
    errorText: state.errorText,
    loadingText: state.loadingText,
    style: resolveEffectiveCardStyle(cardLayoutStyle, previewAppearance),
  };
  const previewResult: IDashboardCardResult = {
    id: 'preview',
    title: previewCardConfig.title,
    aggregate: state.aggregate,
    value: 1234,
    status: 'ready',
  };

  return (
    <Stack tokens={{ childrenGap: 0 }}>
      <Pivot styles={{ root: { marginBottom: 4 } }}>
        <PivotItem headerText="Filtros" itemKey="filtros">
          <Stack tokens={{ childrenGap: 20 }} styles={{ root: { paddingTop: 20 } }}>
            <TextField
              label="Título"
              value={state.title}
              onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) =>
                update({ title: v ?? '' })
              }
              required
              placeholder="Ex: Total de itens"
            />
            <TextField
              label="Subtítulo"
              value={state.subtitle}
              onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) =>
                update({ subtitle: v ?? '' })
              }
              placeholder="Ex: Itens aguardando ação"
            />
            <TextField
              label="Texto quando vazio"
              value={state.emptyValueText}
              onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) =>
                update({ emptyValueText: v ?? '' })
              }
              placeholder="Ex: Nenhum item encontrado"
            />
            <TextField
              label="Texto de erro"
              value={state.errorText}
              onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) =>
                update({ errorText: v ?? '' })
              }
              placeholder="Ex: Falha ao carregar"
            />
            <TextField
              label="Texto de carregamento"
              value={state.loadingText}
              onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) =>
                update({ loadingText: v ?? '' })
              }
              placeholder="Ex: Carregando..."
            />

            <ChoiceGroup
              label="Tipo de agregação"
              options={AGGREGATE_OPTIONS}
              selectedKey={state.aggregate}
              onChange={(
                _: React.FormEvent<HTMLElement | HTMLInputElement> | undefined,
                opt?: IChoiceGroupOption
              ) => {
                if (opt) update({ aggregate: opt.key as TAggregateType });
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
                errorMessage={fieldsLoading ? undefined : listTitle.trim() && !fieldsError && numericFields.length === 0 ? 'Nenhum campo numérico na lista' : undefined}
              />
            )}

            <Separator />

            <Toggle
              label="Aplicar filtro nos dados"
              checked={state.hasFilter}
              onChange={(_: React.MouseEvent<HTMLElement>, checked?: boolean) =>
                update({ hasFilter: !!checked })
              }
              onText="Sim"
              offText="Não"
            />

            {state.hasFilter && (
              <Stack
                tokens={{ childrenGap: 12 }}
                styles={{
                  root: {
                    background: '#faf9f8',
                    padding: '16px',
                    borderRadius: 6,
                    border: '1px solid #edebe9',
                  },
                }}
              >
                {fieldsLoading && (
                  <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                    <Spinner size={SpinnerSize.small} />
                    <Text variant="small">Carregando campos da lista...</Text>
                  </Stack>
                )}
                {fieldsError && (
                  <Text variant="small" styles={{ root: { color: '#d13438' } }}>
                    {fieldsError}
                  </Text>
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
          </Stack>
        </PivotItem>

        <PivotItem headerText="Aparência" itemKey="aparencia">
          <Stack tokens={{ childrenGap: 16 }} styles={{ root: { paddingTop: 20 } }}>
            <Text variant="small" styles={{ root: { color: '#605e5c', marginBottom: 8 } }}>
              Pré-visualização
            </Text>
            <div
              style={{
                background: '#f3f2f1',
                padding: '20px 24px',
                borderRadius: 8,
                display: 'flex',
                justifyContent: 'center',
              }}
            >
              <DashboardCard
                result={previewResult}
                cardConfig={previewCardConfig}
                cardLayoutStyle={cardLayoutStyle}
              />
            </div>

            <Separator />

            <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
              Cores
            </Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Estilo, tipografia e layout são definidos na aba Aparência da visão geral.
            </Text>
            <TextField label="Cor de fundo" value={state.backgroundColor} onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) => update({ backgroundColor: v ?? '' })} placeholder="#ffffff" />
            <TextField label="Cor da borda" value={state.borderColor} onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) => update({ borderColor: v ?? '' })} placeholder="#e2e8f0" />
            <TextField label="Cor do título" value={state.titleColor} onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) => update({ titleColor: v ?? '' })} placeholder="#334155" />
            <TextField label="Cor do subtítulo" value={state.subtitleColor} onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) => update({ subtitleColor: v ?? '' })} placeholder="#64748b" />
            <TextField label="Cor do valor" value={state.valueColor} onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) => update({ valueColor: v ?? '' })} placeholder="#0f172a" />
            {cardLayoutStyle.showIcon && (
              <Dropdown
                label="Ícone"
                options={ICON_OPTIONS}
                selectedKey={state.iconName || ''}
                onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                  if (opt) update({ iconName: String(opt.key) });
                }}
                styles={{ root: { maxWidth: 220 } }}
              />
            )}
          </Stack>
        </PivotItem>
      </Pivot>

      <Separator />

      <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { paddingTop: 16 } }}>
        <PrimaryButton text="Confirmar" onClick={handleConfirm} disabled={!isValid} />
        <DefaultButton text="Voltar" onClick={onBack} />
      </Stack>
    </Stack>
  );
};
