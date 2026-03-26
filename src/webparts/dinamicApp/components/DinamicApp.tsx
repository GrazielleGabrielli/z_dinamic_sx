import * as React from 'react';
import { useState, useCallback, useEffect } from 'react';
import { Text, Stack, Separator, ActionButton } from '@fluentui/react';
import type { IDinamicAppProps } from './IDinamicAppProps';
import { parseConfig } from '../core/config/validators';
import {
  IDashboardCardConfig,
  IChartSeriesConfig,
  IDynamicViewConfig,
  IListViewConfig,
  IListViewFilterConfig,
  IPaginationConfig,
  TChartType,
  TDashboardType,
} from '../core/config/types';
import { effectiveDashboardFilters } from '../core/dashboard/effectiveDashboardFilters';
import { ConfigWizard } from './Wizard/ConfigWizard';
import { DashboardView } from './Dashboard/DashboardView';
import { CardEditorPanel } from './Dashboard/CardEditor/CardEditorPanel';
import { ChartSeriesEditorPanel } from './Dashboard/ChartEditor/ChartSeriesEditorPanel';
import { TableView } from './DataTable/TableView';
import { TableColumnsEditorPanel } from './DataTable/TableColumnsEditorPanel';
import { ProjectManagementView } from './ProjectManagement/ProjectManagementView';

type TDashboardListKey = `card:${string}` | `series:${string}`;

const DinamicApp: React.FC<IDinamicAppProps> = ({ configJson, siteUrl, onSaveConfig }) => {
  const [isEditingWebPart, setIsEditingWebPart] = useState(false);
  const [isEditingCards, setIsEditingCards] = useState(false);
  const [isEditingSeries, setIsEditingSeries] = useState(false);
  const [isEditingTableColumns, setIsEditingTableColumns] = useState(false);
  const [dashboardRefreshKey, setDashboardRefreshKey] = useState(0);
  const [dashboardListSelection, setDashboardListSelection] = useState<{
    key: TDashboardListKey;
    filters: IListViewFilterConfig[];
  } | null>(null);

  const config = parseConfig(configJson ?? undefined);

  useEffect(() => {
    setDashboardListSelection(null);
  }, [config?.dataSource.title]);

  const handleDashboardCardClick = useCallback((card: IDashboardCardConfig) => {
    const filters = effectiveDashboardFilters(card) as IListViewFilterConfig[];
    const key = `card:${card.id}` as TDashboardListKey;
    setDashboardListSelection((prev) => (prev?.key === key ? null : { key, filters }));
  }, []);

  const handleDashboardSeriesClick = useCallback((series: IChartSeriesConfig) => {
    const filters = effectiveDashboardFilters(series) as IListViewFilterConfig[];
    const key = `series:${series.id}` as TDashboardListKey;
    setDashboardListSelection((prev) => (prev?.key === key ? null : { key, filters }));
  }, []);

  const dashKey = dashboardListSelection?.key;
  const selectedCardId =
    dashKey !== undefined && dashKey.indexOf('card:') === 0 ? dashKey.slice('card:'.length) : null;
  const selectedSeriesId =
    dashKey !== undefined && dashKey.indexOf('series:') === 0 ? dashKey.slice('series:'.length) : null;
  const dashboardAppliesListFilter = Boolean(dashboardListSelection?.filters.length);
  const triggerDashboardRefresh = useCallback(() => {
    setDashboardRefreshKey((prev) => prev + 1);
  }, []);

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

  const handleSaveTableColumns = (
    listView: IListViewConfig,
    pagination: IPaginationConfig,
    pdfTemplate?: import('../core/config/types').IPdfTemplateConfig,
    projectManagement?: import('../core/config/types').IProjectManagementConfig
  ): void => {
    if (!config) return;
    onSaveConfig({
      ...config,
      listView,
      pagination,
      ...(projectManagement !== undefined && { projectManagement }),
      ...(pdfTemplate !== undefined && { pdfTemplate }),
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
              refreshKey={dashboardRefreshKey}
              onEditCards={() => setIsEditingCards(true)}
              onEditSeries={() => setIsEditingSeries(true)}
              onCardClick={handleDashboardCardClick}
              selectedCardId={selectedCardId}
              onSeriesClick={handleDashboardSeriesClick}
              selectedSeriesId={selectedSeriesId}
              dashboardAppliesListFilter={dashboardAppliesListFilter}
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
            {config.mode === 'projectManagement' ? 'Editar quadro' : 'Editar colunas'}
          </ActionButton>
        </Stack>

        {config.mode === 'projectManagement' ? (
          <ProjectManagementView
            config={config}
            dashboardListFilters={dashboardListSelection?.filters}
            onItemUpdated={triggerDashboardRefresh}
          />
        ) : (
          <TableView config={config} dashboardListFilters={dashboardListSelection?.filters} />
        )}
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
        mode={config.mode}
        listTitle={config.dataSource.title}
        listView={config.listView}
        pagination={config.pagination}
        projectManagement={config.projectManagement}
        pdfTemplate={config.pdfTemplate}
        onSave={handleSaveTableColumns}
        onDismiss={() => setIsEditingTableColumns(false)}
      />
    </>
  );
};

export default DinamicApp;
