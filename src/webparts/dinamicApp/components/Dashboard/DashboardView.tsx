import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import { Stack, Text, ActionButton, MessageBar, MessageBarType } from '@fluentui/react';
import { FieldsService } from '../../../../services';
import { IDashboardCardConfig, IDashboardConfig, IDataSourceConfig, IChartSeriesConfig } from '../../core/config/types';
import { generateDefaultCards } from '../../core/config/utils';
import { DashboardEngine } from '../../core/dashboard/DashboardEngine';
import { IDashboardCardResult } from '../../core/dashboard/types';
import { DashboardCard } from './DashboardCard';
import { ChartView } from './ChartView';

interface IDashboardViewProps {
  dashboardBlockId: string;
  config: IDashboardConfig;
  dataSource: IDataSourceConfig;
  refreshKey?: number;
  onEditCards: (blockId: string) => void;
  onEditSeries: (blockId: string) => void;
  onSwitchToCharts?: (blockId: string) => void;
  onCardClick?: (card: IDashboardCardConfig, blockId: string) => void;
  selectedCardId?: string | null;
  onSeriesClick?: (series: IChartSeriesConfig, blockId: string) => void;
  selectedSeriesId?: string | null;
  /** Quando há filtros do dashboard aplicados na listagem (para texto auxiliar). */
  dashboardAppliesListFilter?: boolean;
}

export const DashboardView: React.FC<IDashboardViewProps> = ({
  dashboardBlockId,
  config,
  dataSource,
  refreshKey = 0,
  onEditCards,
  onEditSeries,
  onSwitchToCharts,
  onCardClick,
  selectedCardId,
  onSeriesClick,
  selectedSeriesId,
  dashboardAppliesListFilter,
}) => {
  if (config.dashboardType === 'charts') {
    return (
      <ChartView
        config={config}
        dataSource={dataSource}
        refreshKey={refreshKey}
        onEditSeries={() => onEditSeries(dashboardBlockId)}
        onSeriesClick={
          onSeriesClick !== undefined ? (s) => onSeriesClick(s, dashboardBlockId) : undefined
        }
        selectedSeriesId={selectedSeriesId}
        showListFilterHint={dashboardAppliesListFilter === true}
      />
    );
  }

  const engine = React.useMemo(() => new DashboardEngine(), []);
  const fieldsService = useMemo(() => new FieldsService(), []);

  const [results, setResults] = useState<IDashboardCardResult[]>(() =>
    engine.buildLoadingResults(config)
  );
  const [globalError, setGlobalError] = useState<string | undefined>(undefined);
  const [fieldMetadata, setFieldMetadata] = useState<Awaited<ReturnType<FieldsService['getVisibleFields']>> | undefined>(undefined);

  useEffect(() => {
    if (!dataSource.title.trim()) return;
    setFieldMetadata(undefined);
    fieldsService.getVisibleFields(dataSource.title).then(setFieldMetadata).catch(() => setFieldMetadata([]));
  }, [dataSource.title]);

  useEffect(() => {
    setResults(engine.buildLoadingResults(config));
    setGlobalError(undefined);

    const run = fieldMetadata
      ? engine.computeAll(config, dataSource, fieldMetadata)
      : (fieldMetadata === undefined ? Promise.resolve(engine.buildLoadingResults(config)) : engine.computeAll(config, dataSource, []));

    run
      .then((computed) => setResults(computed))
      .catch((err: Error) => setGlobalError(`Erro ao carregar dashboard: ${err.message}`));
  }, [config, dataSource, fieldMetadata, refreshKey]);

  const cardsWithDefaults = useMemo(
    () => (config.cards.length > 0 ? config.cards : generateDefaultCards(config.cardsCount)),
    [config.cards, config.cardsCount]
  );
  const getCardConfig = (id: string) => {
    for (let i = 0; i < cardsWithDefaults.length; i++) {
      if (cardsWithDefaults[i].id === id) return cardsWithDefaults[i];
    }
    return undefined;
  };

  if (results.length === 0) return null;

  return (
    <div style={{ marginBottom: 24 }}>
      <Stack
        horizontal
        horizontalAlign="space-between"
        verticalAlign="center"
        styles={{ root: { marginBottom: 12 } }}
      >
        <Text
          variant="mediumPlus"
          styles={{ root: { fontWeight: 600, color: '#605e5c' } }}
        >
          Dashboard
        </Text>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
          {onSwitchToCharts !== undefined && (
            <ActionButton
              iconProps={{ iconName: 'BarChartVertical' }}
              onClick={() => onSwitchToCharts(dashboardBlockId)}
              styles={{ root: { height: 28, color: '#0078d4' } }}
            >
              Gráficos
            </ActionButton>
          )}
          <ActionButton
            iconProps={{ iconName: 'Edit' }}
            onClick={() => onEditCards(dashboardBlockId)}
            styles={{ root: { height: 28, color: '#0078d4' } }}
          >
            Editar cards
          </ActionButton>
        </Stack>
      </Stack>

      {globalError !== undefined && (
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={false}
          styles={{ root: { marginBottom: 12 } }}
        >
          {globalError}
        </MessageBar>
      )}

      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 16 }}>
        {results.map((result) => {
          const cfg = getCardConfig(result.id);
          return (
            <DashboardCard
              key={result.id}
              result={result}
              cardConfig={cfg}
              selected={selectedCardId === result.id}
              onActivate={onCardClick && cfg ? () => onCardClick(cfg, dashboardBlockId) : undefined}
            />
          );
        })}
      </div>
      {selectedCardId && onCardClick && dashboardAppliesListFilter === true && (
        <Text variant="small" styles={{ root: { color: '#605e5c', marginTop: 12, display: 'block' } }}>
          Filtro do card ativo na listagem — clique de novo no mesmo card para remover.
        </Text>
      )}
    </div>
  );
};
