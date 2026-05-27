import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  IconButton,
  PrimaryButton,
  DefaultButton,
  Separator,
  ChoiceGroup,
  IChoiceGroupOption,
  Pivot,
  PivotItem,
} from '@fluentui/react';
import {
  IDashboardCardConfig,
  IDashboardCardLayoutStyle,
  TChartType,
  TDashboardType,
} from '../../../core/config/types';
import { generateDefaultCards } from '../../../core/config/utils';
import { resolveDashboardCardLayoutStyle, stripLayoutFromCards } from '../../../core/dashboard/utils';
import { ChoiceFieldBreakdownModal } from '../ChoiceFieldBreakdownModal';
import { ChartTypeCard } from '../ChartTypeCard';
import { CardForm } from './CardForm';
import { CardAppearanceTabs } from './CardAppearanceTabs';

const DASHBOARD_TYPE_OPTIONS: IChoiceGroupOption[] = [
  { key: 'cards', text: 'Cards' },
  { key: 'charts', text: 'Gráficos' },
];

const CHART_TYPES: TChartType[] = ['bar', 'line', 'area', 'pie', 'donut'];

export interface ICardEditorSaveOptions {
  dashboardType?: TDashboardType;
  chartType?: TChartType;
  cardLayoutStyle?: IDashboardCardLayoutStyle;
}

interface ICardEditorPanelProps {
  isOpen: boolean;
  listTitle: string;
  listWebServerRelativeUrl?: string;
  cards: IDashboardCardConfig[];
  cardsCount: number;
  dashboardType: TDashboardType;
  chartType?: TChartType;
  cardLayoutStyle?: IDashboardCardLayoutStyle;
  onSave: (cards: IDashboardCardConfig[], options?: ICardEditorSaveOptions) => void;
  onDismiss: () => void;
}

type TPanelView = 'list' | 'form';

const AGGREGATE_LABEL: Record<string, string> = {
  count: 'contagem',
  sum: 'soma',
};

function initLocalCards(
  cards: IDashboardCardConfig[],
  cardsCount: number
): IDashboardCardConfig[] {
  return cards.length > 0 ? [...cards] : generateDefaultCards(cardsCount);
}

function cardFilterSummary(card: IDashboardCardConfig): string | undefined {
  const list =
    card.filters && card.filters.length > 0
      ? card.filters
      : card.filter
        ? [card.filter]
        : [];
  if (list.length === 0) return undefined;
  const f = list[0];
  return `filtro: ${f.field} ${f.operator} "${f.value}"`;
}

export const CardEditorPanel: React.FC<ICardEditorPanelProps> = ({
  isOpen,
  listTitle,
  listWebServerRelativeUrl,
  cards,
  cardsCount,
  dashboardType,
  chartType = 'bar',
  cardLayoutStyle,
  onSave,
  onDismiss,
}) => {
  const [localCards, setLocalCards] = useState<IDashboardCardConfig[]>(() =>
    initLocalCards(cards, cardsCount)
  );
  const [localCardLayoutStyle, setLocalCardLayoutStyle] = useState<IDashboardCardLayoutStyle>(() =>
    cardLayoutStyle ?? resolveDashboardCardLayoutStyle({ cards, cardLayoutStyle })
  );
  const [localDashboardType, setLocalDashboardType] = useState<TDashboardType>(dashboardType);
  const [localChartType, setLocalChartType] = useState<TChartType>(chartType);
  const [view, setView] = useState<TPanelView>('list');
  const [editingIndex, setEditingIndex] = useState<number | undefined>(undefined);
  const [choiceModalOpen, setChoiceModalOpen] = useState(false);

  useEffect(() => {
    if (!isOpen) return;
    setLocalCards(initLocalCards(cards, cardsCount));
    setLocalCardLayoutStyle(
      cardLayoutStyle ?? resolveDashboardCardLayoutStyle({ cards, cardLayoutStyle })
    );
    setLocalDashboardType(dashboardType);
    setLocalChartType(chartType ?? 'bar');
    setView('list');
    setEditingIndex(undefined);
  }, [isOpen, cards, cardsCount, dashboardType, chartType, cardLayoutStyle]);

  const handleEdit = (index: number): void => {
    setEditingIndex(index);
    setView('form');
  };

  const handleAdd = (): void => {
    setEditingIndex(undefined);
    setView('form');
  };

  const handleDelete = (index: number): void => {
    setLocalCards((prev) => prev.filter((_, i) => i !== index));
  };

  const handleConfirmCard = (card: IDashboardCardConfig): void => {
    if (editingIndex !== undefined) {
      setLocalCards((prev) => prev.map((c, i) => (i === editingIndex ? card : c)));
    } else {
      setLocalCards((prev) => [...prev, card]);
    }
    setView('list');
  };

  const handleSave = (): void => {
    onSave(stripLayoutFromCards(localCards), {
      dashboardType: localDashboardType,
      chartType: localChartType,
      cardLayoutStyle: localCardLayoutStyle,
    });
  };

  const getEditingCard = (): IDashboardCardConfig | undefined => {
    if (editingIndex !== undefined && editingIndex < localCards.length) {
      return localCards[editingIndex];
    }
    return undefined;
  };

  const panelHeader =
    view === 'list'
      ? 'Editar cards do dashboard'
      : editingIndex !== undefined
      ? 'Editar card'
      : 'Novo card';

  const renderListView = (): React.ReactElement => (
    <Pivot styles={{ root: { marginTop: 4 } }}>
      <PivotItem headerText="Filtros" itemKey="filtros">
        <Stack tokens={{ childrenGap: 16 }} styles={{ root: { paddingTop: 16 } }}>
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
          {localDashboardType === 'cards' && (
            <>
              <Separator />
              <Stack tokens={{ childrenGap: 0 }}>
                {localCards.length === 0 && (
                  <Text variant="small" styles={{ root: { color: '#a19f9d', padding: '16px 0' } }}>
                    Nenhum card configurado ainda.
                  </Text>
                )}
                {localCards.map((card, index) => {
                  const filterLine = cardFilterSummary(card);
                  return (
                    <React.Fragment key={card.id}>
                      <div
                        style={{
                          padding: '14px 0',
                          display: 'flex',
                          justifyContent: 'space-between',
                          alignItems: 'center',
                        }}
                      >
                        <Stack tokens={{ childrenGap: 2 }}>
                          <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                            {card.title}
                          </Text>
                          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                            {AGGREGATE_LABEL[card.aggregate] ?? card.aggregate}
                            {card.field !== undefined ? ` · campo: ${card.field}` : ''}
                            {filterLine !== undefined ? ` · ${filterLine}` : ''}
                          </Text>
                        </Stack>
                        <Stack horizontal tokens={{ childrenGap: 2 }}>
                          <IconButton
                            iconProps={{ iconName: 'Edit' }}
                            title="Editar"
                            ariaLabel="Editar card"
                            onClick={() => handleEdit(index)}
                            styles={{ root: { color: '#0078d4' } }}
                          />
                          <IconButton
                            iconProps={{ iconName: 'Delete' }}
                            title="Remover"
                            ariaLabel="Remover card"
                            onClick={() => handleDelete(index)}
                            styles={{ root: { color: '#d13438' } }}
                          />
                        </Stack>
                      </div>
                      {index < localCards.length - 1 && (
                        <Separator styles={{ root: { padding: 0 } }} />
                      )}
                    </React.Fragment>
                  );
                })}
                <Stack
                  horizontal
                  tokens={{ childrenGap: 8 }}
                  styles={{ root: { marginTop: 20, flexWrap: 'wrap' } }}
                >
                  <DefaultButton
                    iconProps={{ iconName: 'Add' }}
                    text="Adicionar card"
                    onClick={handleAdd}
                  />
                  <DefaultButton
                    iconProps={{ iconName: 'LightningBolt' }}
                    text="Avançada"
                    onClick={() => setChoiceModalOpen(true)}
                  />
                </Stack>
              </Stack>
            </>
          )}
        </Stack>
      </PivotItem>
      {localDashboardType === 'cards' && (
        <PivotItem headerText="Aparência" itemKey="aparencia">
          <Stack styles={{ root: { paddingTop: 16 } }}>
            <CardAppearanceTabs value={localCardLayoutStyle} onChange={setLocalCardLayoutStyle} />
          </Stack>
        </PivotItem>
      )}
    </Pivot>
  );

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.custom}
      customWidth="72vw"
      styles={{
        main: {
          width: 'min(72vw, calc(100vw - 16px))',
          maxWidth: 'min(72vw, calc(100vw - 16px))',
          borderRadius: '12px 0 0 12px',
          boxShadow: '-8px 0 40px rgba(0, 0, 0, 0.12)',
        },
        scrollableContent: { overflowX: 'hidden', background: '#f5f4f2' },
        content: { overflowX: 'hidden', minWidth: 0, background: '#f5f4f2' },
        header: {
          background: '#f8f7f6',
          borderBottom: '1px solid #e8e6e4',
          flexShrink: 0,
        },
        footer: { borderTop: '1px solid #e8e6e4', background: '#faf9f8', paddingTop: 12 },
      }}
      headerText={panelHeader}
      closeButtonAriaLabel="Fechar"
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <Stack
          horizontal
          tokens={{ childrenGap: 8 }}
          styles={{ root: { paddingBottom: 16 } }}
        >
          {view === 'list' && (
            <>
              <PrimaryButton text="Salvar" onClick={handleSave} />
              <DefaultButton text="Cancelar" onClick={onDismiss} />
            </>
          )}
        </Stack>
      )}
    >
      <ChoiceFieldBreakdownModal
        isOpen={choiceModalOpen}
        onDismiss={() => setChoiceModalOpen(false)}
        listTitle={listTitle}
        listWebServerRelativeUrl={listWebServerRelativeUrl}
        target="cards"
        onApply={(items, mergeMode) => {
          const next = items as IDashboardCardConfig[];
          setLocalCards((prev) => (mergeMode === 'replace' ? next : [...prev, ...next]));
        }}
      />
      <div style={{ paddingTop: 16, minWidth: 0, maxWidth: '100%', boxSizing: 'border-box' }}>
        {view === 'list' && renderListView()}
        {view === 'form' && (
          <CardForm
            listTitle={listTitle}
            listWebServerRelativeUrl={listWebServerRelativeUrl}
            card={getEditingCard()}
            cardLayoutStyle={localCardLayoutStyle}
            onConfirm={handleConfirmCard}
            onBack={() => setView('list')}
          />
        )}
      </div>
    </Panel>
  );
};
