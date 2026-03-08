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
} from '@fluentui/react';
import {
  IDashboardCardConfig,
  IDashboardCardStyleConfig,
  TAggregateType,
  TFilterOperator,
  TCardVariant,
  TBorderRadius,
  TPadding,
  TShadow,
  TTitleSize,
  TSubtitleSize,
  TValueSize,
  TFontWeight,
  TAlign,
  TIconPosition,
  TLoadingStyle,
} from '../../../core/config/types';
import { mergeWithDefaultStyle } from '../../../core/dashboard/utils/dashboardCardStyles';
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
  card: IDashboardCardConfig | undefined;
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
  filterField: string;
  filterOperator: TFilterOperator;
  filterValue: string;
  variant: TCardVariant;
  borderRadius: TBorderRadius;
  padding: TPadding;
  shadow: TShadow;
  border: boolean;
  backgroundColor: string;
  borderColor: string;
  titleColor: string;
  subtitleColor: string;
  valueColor: string;
  titleSize: TTitleSize;
  subtitleSize: TSubtitleSize;
  valueSize: TValueSize;
  titleWeight: TFontWeight;
  valueWeight: TFontWeight;
  align: TAlign;
  showSubtitle: boolean;
  showValue: boolean;
  showIcon: boolean;
  iconName: string;
  iconPosition: TIconPosition;
  loadingStyle: TLoadingStyle;
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

const VARIANT_OPTIONS: IDropdownOption[] = [
  { key: 'default', text: 'Padrão' },
  { key: 'outlined', text: 'Contorno' },
  { key: 'soft', text: 'Suave' },
  { key: 'solid', text: 'Sólido' },
];

const BORDER_RADIUS_OPTIONS: IDropdownOption[] = [
  { key: 'none', text: 'Nenhum' },
  { key: 'sm', text: 'Pequeno' },
  { key: 'md', text: 'Médio' },
  { key: 'lg', text: 'Grande' },
  { key: 'xl', text: 'Extra grande' },
  { key: 'full', text: 'Total' },
];

const PADDING_OPTIONS: IDropdownOption[] = [
  { key: 'sm', text: 'Pequeno' },
  { key: 'md', text: 'Médio' },
  { key: 'lg', text: 'Grande' },
];

const SHADOW_OPTIONS: IDropdownOption[] = [
  { key: 'none', text: 'Nenhum' },
  { key: 'sm', text: 'Pequeno' },
  { key: 'md', text: 'Médio' },
  { key: 'lg', text: 'Grande' },
];

const TITLE_SIZE_OPTIONS: IDropdownOption[] = [
  { key: 'xs', text: 'Extra pequeno' },
  { key: 'sm', text: 'Pequeno' },
  { key: 'md', text: 'Médio' },
  { key: 'lg', text: 'Grande' },
];

const SUBTITLE_SIZE_OPTIONS: IDropdownOption[] = [
  { key: 'xs', text: 'Extra pequeno' },
  { key: 'sm', text: 'Pequeno' },
  { key: 'md', text: 'Médio' },
];

const VALUE_SIZE_OPTIONS: IDropdownOption[] = [
  { key: 'lg', text: 'Grande' },
  { key: 'xl', text: 'Extra grande' },
  { key: '2xl', text: '2x grande' },
  { key: '3xl', text: '3x grande' },
];

const FONT_WEIGHT_OPTIONS: IDropdownOption[] = [
  { key: 'normal', text: 'Normal' },
  { key: 'medium', text: 'Médio' },
  { key: 'semibold', text: 'Semi-negrito' },
  { key: 'bold', text: 'Negrito' },
];

const ALIGN_OPTIONS: IDropdownOption[] = [
  { key: 'left', text: 'Esquerda' },
  { key: 'center', text: 'Centro' },
  { key: 'right', text: 'Direita' },
];

const ICON_POSITION_OPTIONS: IDropdownOption[] = [
  { key: 'left', text: 'Esquerda' },
  { key: 'top', text: 'Topo' },
  { key: 'right', text: 'Direita' },
];

const LOADING_STYLE_OPTIONS: IDropdownOption[] = [
  { key: 'skeleton', text: 'Skeleton' },
  { key: 'spinner', text: 'Spinner' },
  { key: 'text', text: 'Texto' },
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

function initState(card?: IDashboardCardConfig): ICardFormState {
  const s = mergeWithDefaultStyle(card?.style);
  return {
    title: card?.title ?? '',
    subtitle: card?.subtitle ?? '',
    emptyValueText: card?.emptyValueText ?? 'Nenhum item',
    errorText: card?.errorText ?? 'Erro ao carregar',
    loadingText: card?.loadingText ?? 'Carregando...',
    aggregate: card?.aggregate ?? 'count',
    field: card?.field ?? '',
    hasFilter: card?.filter !== undefined,
    filterField: card?.filter?.field ?? '',
    filterOperator: card?.filter?.operator ?? 'eq',
    filterValue: card?.filter?.value ?? '',
    variant: s.variant,
    borderRadius: s.borderRadius,
    padding: s.padding,
    shadow: s.shadow,
    border: s.border,
    backgroundColor: s.backgroundColor ?? '',
    borderColor: s.borderColor ?? '',
    titleColor: s.titleColor ?? '',
    subtitleColor: s.subtitleColor ?? '',
    valueColor: s.valueColor ?? '',
    titleSize: s.titleSize,
    subtitleSize: s.subtitleSize,
    valueSize: s.valueSize,
    titleWeight: s.titleWeight,
    valueWeight: s.valueWeight,
    align: s.align,
    showSubtitle: s.showSubtitle,
    showValue: s.showValue,
    showIcon: s.showIcon,
    iconName: s.iconName ?? '',
    iconPosition: s.iconPosition,
    loadingStyle: s.loadingStyle,
  };
}

function buildCardStyle(state: ICardFormState): IDashboardCardStyleConfig {
  const style: IDashboardCardStyleConfig = {
    variant: state.variant,
    borderRadius: state.borderRadius,
    padding: state.padding,
    shadow: state.shadow,
    border: state.border,
    titleSize: state.titleSize,
    subtitleSize: state.subtitleSize,
    valueSize: state.valueSize,
    titleWeight: state.titleWeight,
    valueWeight: state.valueWeight,
    align: state.align,
    showIcon: state.showIcon,
    iconPosition: state.iconPosition,
    showSubtitle: state.showSubtitle,
    showValue: state.showValue,
    loadingStyle: state.loadingStyle,
  };
  if (state.backgroundColor) style.backgroundColor = state.backgroundColor;
  if (state.borderColor) style.borderColor = state.borderColor;
  if (state.titleColor) style.titleColor = state.titleColor;
  if (state.subtitleColor) style.subtitleColor = state.subtitleColor;
  if (state.valueColor) style.valueColor = state.valueColor;
  if (state.showIcon && state.iconName) style.iconName = state.iconName;
  return style;
}

function buildCard(state: ICardFormState, existingId?: string): IDashboardCardConfig {
  const card: IDashboardCardConfig = {
    id: existingId ?? `card_${String(Date.now())}`,
    title: state.title.trim(),
    aggregate: state.aggregate,
    style: buildCardStyle(state),
  };
  if (state.subtitle.trim()) card.subtitle = state.subtitle.trim();
  if (state.emptyValueText.trim()) card.emptyValueText = state.emptyValueText.trim();
  if (state.errorText.trim()) card.errorText = state.errorText.trim();
  if (state.loadingText.trim()) card.loadingText = state.loadingText.trim();
  if (state.aggregate === 'sum' && state.field.trim().length > 0) {
    card.field = state.field.trim();
  }
  if (state.hasFilter && state.filterField.trim().length > 0 && state.filterValue.trim().length > 0) {
    card.filter = {
      field: state.filterField.trim(),
      operator: state.filterOperator,
      value: state.filterValue.trim(),
    };
  }
  return card;
}

export const CardForm: React.FC<ICardFormProps> = ({ listTitle, card, onConfirm, onBack }) => {
  const [state, setState] = useState<ICardFormState>(() => initState(card));
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
      .getVisibleFields(listTitle.trim())
      .then((fields) => {
        setListFields(fields);
        setFieldsLoading(false);
      })
      .catch((err) => {
        setListFields([]);
        setFieldsError(err instanceof Error ? err.message : String(err));
        setFieldsLoading(false);
      });
  }, [listTitle]);

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
    (!state.showIcon || state.iconName.trim().length > 0);

  const handleConfirm = (): void => {
    if (!isValid) return;
    onConfirm(buildCard(state, card?.id));
  };

  const previewCardConfig: IDashboardCardConfig = {
    id: 'preview',
    title: state.title.trim() || 'Título do card',
    subtitle: state.subtitle.trim() || undefined,
    aggregate: state.aggregate,
    emptyValueText: state.emptyValueText,
    errorText: state.errorText,
    loadingText: state.loadingText,
    style: buildCardStyle(state),
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
        <PivotItem headerText="Dados">
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
          </Stack>
        </PivotItem>

        <PivotItem headerText="Aparência">
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
              <DashboardCard result={previewResult} cardConfig={previewCardConfig} />
            </div>

            <Separator />

            <Dropdown
              label="Variante"
              options={VARIANT_OPTIONS}
              selectedKey={state.variant}
              onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                if (opt) update({ variant: opt.key as TCardVariant });
              }}
              styles={{ root: { maxWidth: 200 } }}
            />
            <Dropdown
              label="Borda (cantos)"
              options={BORDER_RADIUS_OPTIONS}
              selectedKey={state.borderRadius}
              onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                if (opt) update({ borderRadius: opt.key as TBorderRadius });
              }}
              styles={{ root: { maxWidth: 200 } }}
            />
            <Dropdown
              label="Preenchimento"
              options={PADDING_OPTIONS}
              selectedKey={state.padding}
              onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                if (opt) update({ padding: opt.key as TPadding });
              }}
              styles={{ root: { maxWidth: 200 } }}
            />
            <Dropdown
              label="Sombra"
              options={SHADOW_OPTIONS}
              selectedKey={state.shadow}
              onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                if (opt) update({ shadow: opt.key as TShadow });
              }}
              styles={{ root: { maxWidth: 200 } }}
            />
            <Toggle
              label="Exibir borda"
              checked={state.border}
              onChange={(_: React.MouseEvent<HTMLElement>, checked?: boolean) =>
                update({ border: !!checked })
              }
            />

            <TextField label="Cor de fundo" value={state.backgroundColor} onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) => update({ backgroundColor: v ?? '' })} placeholder="#ffffff" />
            <TextField label="Cor da borda" value={state.borderColor} onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) => update({ borderColor: v ?? '' })} placeholder="#e2e8f0" />
            <TextField label="Cor do título" value={state.titleColor} onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) => update({ titleColor: v ?? '' })} placeholder="#334155" />
            <TextField label="Cor do subtítulo" value={state.subtitleColor} onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) => update({ subtitleColor: v ?? '' })} placeholder="#64748b" />
            <TextField label="Cor do valor" value={state.valueColor} onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) => update({ valueColor: v ?? '' })} placeholder="#0f172a" />
          </Stack>
        </PivotItem>

        <PivotItem headerText="Tipografia">
          <Stack tokens={{ childrenGap: 16 }} styles={{ root: { paddingTop: 20 } }}>
            <Dropdown
              label="Tamanho do título"
              options={TITLE_SIZE_OPTIONS}
              selectedKey={state.titleSize}
              onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                if (opt) update({ titleSize: opt.key as TTitleSize });
              }}
              styles={{ root: { maxWidth: 200 } }}
            />
            <Dropdown
              label="Tamanho do subtítulo"
              options={SUBTITLE_SIZE_OPTIONS}
              selectedKey={state.subtitleSize}
              onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                if (opt) update({ subtitleSize: opt.key as TSubtitleSize });
              }}
              styles={{ root: { maxWidth: 200 } }}
            />
            <Dropdown
              label="Tamanho do valor"
              options={VALUE_SIZE_OPTIONS}
              selectedKey={state.valueSize}
              onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                if (opt) update({ valueSize: opt.key as TValueSize });
              }}
              styles={{ root: { maxWidth: 200 } }}
            />
            <Dropdown
              label="Peso do título"
              options={FONT_WEIGHT_OPTIONS}
              selectedKey={state.titleWeight}
              onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                if (opt) update({ titleWeight: opt.key as TFontWeight });
              }}
              styles={{ root: { maxWidth: 200 } }}
            />
            <Dropdown
              label="Peso do valor"
              options={FONT_WEIGHT_OPTIONS}
              selectedKey={state.valueWeight}
              onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                if (opt) update({ valueWeight: opt.key as TFontWeight });
              }}
              styles={{ root: { maxWidth: 200 } }}
            />
          </Stack>
        </PivotItem>

        <PivotItem headerText="Layout">
          <Stack tokens={{ childrenGap: 16 }} styles={{ root: { paddingTop: 20 } }}>
            <Dropdown
              label="Alinhamento"
              options={ALIGN_OPTIONS}
              selectedKey={state.align}
              onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                if (opt) update({ align: opt.key as TAlign });
              }}
              styles={{ root: { maxWidth: 200 } }}
            />
            <Toggle
              label="Exibir subtítulo"
              checked={state.showSubtitle}
              onChange={(_: React.MouseEvent<HTMLElement>, checked?: boolean) =>
                update({ showSubtitle: !!checked })
              }
            />
            <Toggle
              label="Exibir valor"
              checked={state.showValue}
              onChange={(_: React.MouseEvent<HTMLElement>, checked?: boolean) =>
                update({ showValue: !!checked })
              }
            />
            <Toggle
              label="Exibir ícone"
              checked={state.showIcon}
              onChange={(_: React.MouseEvent<HTMLElement>, checked?: boolean) =>
                update({ showIcon: !!checked })
              }
            />
            {state.showIcon && (
              <>
                <Dropdown
                  label="Ícone"
                  options={ICON_OPTIONS}
                  selectedKey={state.iconName || ''}
                  onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                    if (opt) update({ iconName: String(opt.key) });
                  }}
                  styles={{ root: { maxWidth: 220 } }}
                />
                <Dropdown
                  label="Posição do ícone"
                  options={ICON_POSITION_OPTIONS}
                  selectedKey={state.iconPosition}
                  onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                    if (opt) update({ iconPosition: opt.key as TIconPosition });
                  }}
                  styles={{ root: { maxWidth: 200 } }}
                />
              </>
            )}
            <Dropdown
              label="Estilo de carregamento"
              options={LOADING_STYLE_OPTIONS}
              selectedKey={state.loadingStyle}
              onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                if (opt) update({ loadingStyle: opt.key as TLoadingStyle });
              }}
              styles={{ root: { maxWidth: 200 } }}
            />
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
