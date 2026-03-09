import * as React from 'react';
import { useState } from 'react';
import { Text, Stack, Separator, ActionButton } from '@fluentui/react';
import type { IDinamicAppProps } from './IDinamicAppProps';
import { parseConfig } from '../core/config/validators';
import { IDashboardCardConfig, IChartSeriesConfig, IDynamicViewConfig } from '../core/config/types';
import { ConfigWizard } from './Wizard/ConfigWizard';
import { DashboardView } from './Dashboard/DashboardView';
import { CardEditorPanel } from './Dashboard/CardEditor/CardEditorPanel';
import { ChartSeriesEditorPanel } from './Dashboard/ChartEditor/ChartSeriesEditorPanel';
import { ListView } from './ListView/ListView';

const DinamicApp: React.FC<IDinamicAppProps> = ({ configJson, siteUrl, onSaveConfig }) => {
  const [isEditingWebPart, setIsEditingWebPart] = useState(false);
  const [isEditingCards, setIsEditingCards] = useState(false);
  const [isEditingSeries, setIsEditingSeries] = useState(false);

  const config = parseConfig(configJson ?? undefined);

  const handleWizardComplete = (newConfig: IDynamicViewConfig): void => {
    onSaveConfig(newConfig);
    setIsEditingWebPart(false);
  };

  const handleSaveCards = (cards: IDashboardCardConfig[]): void => {
    if (!config) return;
    onSaveConfig({
      ...config,
      dashboard: {
        ...config.dashboard,
        cards,
        cardsCount: cards.length,
      },
    });
    setIsEditingCards(false);
  };

  const handleSaveSeries = (chartSeries: IChartSeriesConfig[]): void => {
    if (!config) return;
    onSaveConfig({
      ...config,
      dashboard: {
        ...config.dashboard,
        chartSeries,
      },
    });
    setIsEditingSeries(false);
  };

  if (config === undefined || isEditingWebPart) {
    return (
      <ConfigWizard
        siteUrl={siteUrl}
        onComplete={handleWizardComplete}
        initialValues={config}
        onCancel={config !== undefined ? () => setIsEditingWebPart(false) : undefined}
      />
    );
  }

  const showDashboard =
    config.dashboard.enabled &&
    (config.dashboard.dashboardType === 'charts' || config.dashboard.cardsCount > 0);

  return (
    <>
      {/* Toolbar */}
      <div
        style={{
          display: 'flex',
          justifyContent: 'flex-end',
          padding: '6px 16px 0',
          borderBottom: '1px solid #f3f2f1',
        }}
      >
        <ActionButton
          iconProps={{ iconName: 'Settings' }}
          onClick={() => setIsEditingWebPart(true)}
          styles={{ root: { color: '#605e5c', fontSize: 12 } }}
        >
          Editar configuração
        </ActionButton>
      </div>

      <Stack styles={{ root: { padding: '20px 24px 0' } }}>
        {showDashboard && (
          <>
            <DashboardView
              config={config.dashboard}
              dataSource={config.dataSource}
              onEditCards={() => setIsEditingCards(true)}
              onEditSeries={() => setIsEditingSeries(true)}
            />
            <Separator />
          </>
        )}

        <Stack tokens={{ childrenGap: 6 }} styles={{ root: { padding: '16px 0' } }}>
          <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
            {config.dataSource.title}
          </Text>
          <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
            Modo: {config.mode} · Origem: {config.dataSource.kind}
          </Text>
        </Stack>

        <ListView config={config} />
      </Stack>

      <CardEditorPanel
        isOpen={isEditingCards}
        listTitle={config.dataSource.title}
        cards={config.dashboard.cards}
        cardsCount={config.dashboard.cardsCount}
        onSave={handleSaveCards}
        onDismiss={() => setIsEditingCards(false)}
      />

      <ChartSeriesEditorPanel
        isOpen={isEditingSeries}
        listTitle={config.dataSource.title}
        series={config.dashboard.chartSeries ?? []}
        onSave={handleSaveSeries}
        onDismiss={() => setIsEditingSeries(false)}
      />
    </>
  );
};

export default DinamicApp;
