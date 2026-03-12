import * as React from 'react';
import { useState } from 'react';
import { Text, Stack, Separator, ActionButton } from '@fluentui/react';
import type { IDinamicAppProps } from './IDinamicAppProps';
import { parseConfig } from '../core/config/validators';
import { IDashboardCardConfig, IChartSeriesConfig, IDynamicViewConfig, IListViewConfig, IPaginationConfig, TChartType, TDashboardType } from '../core/config/types';
import { ConfigWizard } from './Wizard/ConfigWizard';
import { DashboardView } from './Dashboard/DashboardView';
import { CardEditorPanel } from './Dashboard/CardEditor/CardEditorPanel';
import { ChartSeriesEditorPanel } from './Dashboard/ChartEditor/ChartSeriesEditorPanel';
import { TableView } from './DataTable/TableView';
import { TableColumnsEditorPanel } from './DataTable/TableColumnsEditorPanel';

const DinamicApp: React.FC<IDinamicAppProps> = ({ configJson, siteUrl, onSaveConfig }) => {
  const [isEditingWebPart, setIsEditingWebPart] = useState(false);
  const [isEditingCards, setIsEditingCards] = useState(false);
  const [isEditingSeries, setIsEditingSeries] = useState(false);
  const [isEditingTableColumns, setIsEditingTableColumns] = useState(false);

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

  const handleSaveSeries = (
    chartSeries: IChartSeriesConfig[],
    options?: { dashboardType?: TDashboardType; chartType?: TChartType }
  ): void => {
    if (!config) return;
    onSaveConfig({
      ...config,
      dashboard: {
        ...config.dashboard,
        chartSeries,
        ...(options?.dashboardType !== undefined && { dashboardType: options.dashboardType }),
        ...(options?.chartType !== undefined && { chartType: options.chartType }),
      },
    });
    setIsEditingSeries(false);
  };

  const handleSaveTableColumns = (listView: IListViewConfig, pagination: IPaginationConfig): void => {
    if (!config) return;
    onSaveConfig({
      ...config,
      listView,
      pagination,
    });
    setIsEditingTableColumns(false);
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

        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          tokens={{ childrenGap: 8 }}
          styles={{ root: { padding: '16px 0 8px' } }}
        >
          <Stack tokens={{ childrenGap: 6 }}>
            <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
              {config.dataSource.title}
            </Text>
            <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
              Modo: {config.mode} · Origem: {config.dataSource.kind}
            </Text>
          </Stack>
          <ActionButton
            iconProps={{ iconName: 'ColumnOptions' }}
            onClick={() => setIsEditingTableColumns(true)}
            styles={{ root: { height: 28, color: '#0078d4' } }}
          >
            Editar colunas
          </ActionButton>
        </Stack>

        <TableView config={config} />
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
        dashboardType={config.dashboard.dashboardType}
        chartType={config.dashboard.chartType}
        onSave={handleSaveSeries}
        onDismiss={() => setIsEditingSeries(false)}
      />

      <TableColumnsEditorPanel
        isOpen={isEditingTableColumns}
        listTitle={config.dataSource.title}
        listView={config.listView}
        pagination={config.pagination}
        onSave={handleSaveTableColumns}
        onDismiss={() => setIsEditingTableColumns(false)}
      />
    </>
  );
};

export default DinamicApp;
