import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import { Stack, Text, ActionButton, MessageBar, MessageBarType } from '@fluentui/react';
import { IDashboardConfig, IDataSourceConfig } from '../../core/config/types';
import { generateDefaultCards } from '../../core/config/utils';
import { DashboardEngine } from '../../core/dashboard/DashboardEngine';
import { IDashboardCardResult } from '../../core/dashboard/types';
import { DashboardCard } from './DashboardCard';
import { ChartView } from './ChartView';

interface IDashboardViewProps {
  config: IDashboardConfig;
  dataSource: IDataSourceConfig;
  onEditCards: () => void;
  onEditSeries: () => void;
}

export const DashboardView: React.FC<IDashboardViewProps> = ({
  config,
  dataSource,
  onEditCards,
  onEditSeries,
}) => {
  if (config.dashboardType === 'charts') {
    return <ChartView config={config} dataSource={dataSource} onEditSeries={onEditSeries} />;
  }

  const engine = React.useMemo(() => new DashboardEngine(), []);

  const [results, setResults] = useState<IDashboardCardResult[]>(() =>
    engine.buildLoadingResults(config)
  );
  const [globalError, setGlobalError] = useState<string | undefined>(undefined);

  useEffect(() => {
    setResults(engine.buildLoadingResults(config));
    setGlobalError(undefined);

    engine
      .computeAll(config, dataSource)
      .then((computed) => {
        setResults(computed);
      })
      .catch((err: Error) => {
        setGlobalError(`Erro ao carregar dashboard: ${err.message}`);
      });
  }, [config, dataSource]);

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
        <ActionButton
          iconProps={{ iconName: 'Edit' }}
          onClick={onEditCards}
          styles={{ root: { height: 28, color: '#0078d4' } }}
        >
          Editar cards
        </ActionButton>
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
        {results.map((result) => (
          <DashboardCard key={result.id} result={result} cardConfig={getCardConfig(result.id)} />
        ))}
      </div>
    </div>
  );
};
