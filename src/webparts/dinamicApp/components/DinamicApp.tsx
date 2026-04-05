import * as React from 'react';
import { useState, useCallback, useEffect } from 'react';
import { Text, Stack, ActionButton } from '@fluentui/react';
import type { IDinamicAppProps } from './IDinamicAppProps';
import { parseConfig } from '../core/config/validators';
import {
  IDashboardCardConfig,
  IDashboardConfig,
  IChartSeriesConfig,
  IDynamicViewConfig,
  IListPageBlock,
  IListPageLayoutConfig,
  IListViewConfig,
  IListViewFilterConfig,
  IPaginationConfig,
  TChartType,
  TDashboardType,
} from '../core/config/types';
import {
  defaultListPageLayoutFromLegacy,
  findListPageBlockById,
  getDashboardForEditor,
  getEffectiveListPageSections,
  LEGACY_LIST_PAGE_DASHBOARD_BLOCK_ID,
  replaceBlockInListPageLayout,
  saveDashboardForListBlock,
} from '../core/listPage/listPageLayoutUtils';
import { effectiveDashboardFilters } from '../core/dashboard/effectiveDashboardFilters';
import {
  chartSeriesToDashboardCards,
  dashboardCardsToChartSeries,
} from '../core/dashboard/chartSeriesToDashboardCards';
import { generateDefaultCards } from '../core/config/utils';
import { ConfigWizard } from './Wizard/ConfigWizard';
import { CardEditorPanel } from './Dashboard/CardEditor/CardEditorPanel';
import { ChartSeriesEditorPanel } from './Dashboard/ChartEditor/ChartSeriesEditorPanel';
import { TableColumnsEditorPanel } from './DataTable/TableColumnsEditorPanel';
import { ProjectManagementView } from './ProjectManagement/ProjectManagementView';
import { ListPageRenderer, type TListPageDashboardListSelection } from './ListPage/ListPageRenderer';
import { ListPageLayoutEditorPanel } from './ListPage/ListPageLayoutEditorPanel';
import { ListPageBlockConfigPanel } from './ListPage/ListPageBlockConfigPanel';

const DinamicApp: React.FC<IDinamicAppProps> = ({ configJson, siteUrl, instanceScopeId, onSaveConfig }) => {
  const [isEditingWebPart, setIsEditingWebPart] = useState(false);
  const [isEditingCards, setIsEditingCards] = useState(false);
  const [isEditingSeries, setIsEditingSeries] = useState(false);
  const [isEditingTableColumns, setIsEditingTableColumns] = useState(false);
  const [isEditingPageLayout, setIsEditingPageLayout] = useState(false);
  const [listPageContentBlockId, setListPageContentBlockId] = useState<string | null>(null);
  const [editingDashboardBlockId, setEditingDashboardBlockId] = useState<string | null>(null);
  const [dashboardRefreshKey, setDashboardRefreshKey] = useState(0);
  const [dashboardListSelection, setDashboardListSelection] =
    useState<TListPageDashboardListSelection | null>(null);

  const config = parseConfig(configJson ?? undefined);

  useEffect(() => {
    setDashboardListSelection(null);
  }, [config?.dataSource.title]);

  const handleDashboardCardClick = useCallback((card: IDashboardCardConfig, blockId: string) => {
    const filters = effectiveDashboardFilters(card) as IListViewFilterConfig[];
    setDashboardListSelection((prev) =>
      prev !== null &&
      prev.kind === 'card' &&
      prev.blockId === blockId &&
      prev.entityId === card.id
        ? null
        : { blockId, kind: 'card', entityId: card.id, filters }
    );
  }, []);

  const handleDashboardSeriesClick = useCallback((series: IChartSeriesConfig, blockId: string) => {
    const filters = effectiveDashboardFilters(series) as IListViewFilterConfig[];
    setDashboardListSelection((prev) =>
      prev !== null &&
      prev.kind === 'series' &&
      prev.blockId === blockId &&
      prev.entityId === series.id
        ? null
        : { blockId, kind: 'series', entityId: series.id, filters }
    );
  }, []);

  const dashboardAppliesListFilter = Boolean(dashboardListSelection?.filters.length);
  const triggerDashboardRefresh = useCallback(() => {
    setDashboardRefreshKey((prev) => prev + 1);
  }, []);

  const handleWizardComplete = (newConfig: IDynamicViewConfig): void => {
    onSaveConfig(newConfig);
    setIsEditingWebPart(false);
  };

  const handleSwitchDashboardToCharts = useCallback((blockId: string): void => {
    if (!config) return;
    const dash = getDashboardForEditor(config, blockId);
    const cards =
      dash.cards.length > 0 ? dash.cards : generateDefaultCards(dash.cardsCount);
    const chartSeries = dashboardCardsToChartSeries(cards);
    setDashboardListSelection(null);
    const next: IDashboardConfig = {
      ...dash,
      dashboardType: 'charts',
      chartSeries,
      chartType: dash.chartType ?? 'bar',
    };
    onSaveConfig(saveDashboardForListBlock(config, blockId, next));
  }, [config, onSaveConfig]);

  const handleSaveCards = (
    cards: IDashboardCardConfig[],
    options?: { dashboardType?: TDashboardType; chartType?: TChartType }
  ): void => {
    if (!config) return;
    const bid = editingDashboardBlockId ?? LEGACY_LIST_PAGE_DASHBOARD_BLOCK_ID;
    const source = getDashboardForEditor(config, bid);
    const nextType = options?.dashboardType ?? source.dashboardType;
    const baseDash = {
      ...source,
      cards,
      cardsCount: cards.length,
      ...(options?.dashboardType !== undefined && { dashboardType: options.dashboardType }),
      ...(options?.chartType !== undefined && { chartType: options.chartType }),
    };
    const dashboard =
      nextType === 'charts'
        ? {
            ...baseDash,
            dashboardType: 'charts' as const,
            chartSeries: dashboardCardsToChartSeries(cards),
          }
        : baseDash;
    if (nextType === 'charts') setDashboardListSelection(null);
    onSaveConfig(saveDashboardForListBlock(config, bid, dashboard));
    setIsEditingCards(false);
    setEditingDashboardBlockId(null);
  };

  const handleSaveSeries = (
    chartSeries: IChartSeriesConfig[],
    options?: { dashboardType?: TDashboardType; chartType?: TChartType }
  ): void => {
    if (!config) return;
    const bid = editingDashboardBlockId ?? LEGACY_LIST_PAGE_DASHBOARD_BLOCK_ID;
    const source = getDashboardForEditor(config, bid);
    const nextType = options?.dashboardType ?? source.dashboardType;
    const baseDash = {
      ...source,
      chartSeries,
      ...(options?.dashboardType !== undefined && { dashboardType: options.dashboardType }),
      ...(options?.chartType !== undefined && { chartType: options.chartType }),
    };
    const dashboard =
      nextType === 'cards'
        ? {
            ...baseDash,
            cards: chartSeriesToDashboardCards(chartSeries),
            cardsCount: chartSeries.length,
          }
        : baseDash;
    onSaveConfig(saveDashboardForListBlock(config, bid, dashboard));
    setIsEditingSeries(false);
    setEditingDashboardBlockId(null);
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

  const handleSaveListPageLayout = (layout: IListPageLayoutConfig): void => {
    if (!config) return;
    onSaveConfig({
      ...config,
      listPageLayout: layout,
    });
    setIsEditingPageLayout(false);
  };

  const handleApplyListContentBlock = (next: IListPageBlock): void => {
    if (!config?.listPageLayout) return;
    onSaveConfig({
      ...config,
      listPageLayout: replaceBlockInListPageLayout(config.listPageLayout, next.id, next),
    });
    setListPageContentBlockId(null);
  };

  const editingListContentBlock =
    listPageContentBlockId !== null && config?.listPageLayout
      ? findListPageBlockById(config.listPageLayout, listPageContentBlockId)
      : null;
  const listContentBlockPanelOpen =
    editingListContentBlock !== null &&
    (editingListContentBlock.type === 'banner' ||
      editingListContentBlock.type === 'editor' ||
      editingListContentBlock.type === 'sectionTitle' ||
      editingListContentBlock.type === 'alert');

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

  const activeListDashboard = getDashboardForEditor(config, editingDashboardBlockId);

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
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }} wrap>
            {config.mode === 'list' && (
              <ActionButton
                iconProps={{ iconName: 'TripleColumn' }}
                onClick={() => setIsEditingPageLayout(true)}
                styles={{ root: { height: 28, color: '#0078d4' } }}
              >
                Layout da página
              </ActionButton>
            )}
            <ActionButton
              iconProps={{ iconName: 'ColumnOptions' }}
              onClick={() => setIsEditingTableColumns(true)}
              styles={{ root: { height: 28, color: '#0078d4' } }}
            >
              {config.mode === 'projectManagement' ? 'Editar quadro' : 'Editar colunas'}
            </ActionButton>
          </Stack>
        </Stack>

        {config.mode === 'projectManagement' ? (
          <ProjectManagementView
            config={config}
            dashboardListFilters={dashboardListSelection?.filters}
            onItemUpdated={triggerDashboardRefresh}
          />
        ) : (
          <ListPageRenderer
            config={config}
            sections={getEffectiveListPageSections(config)}
            instanceScopeId={instanceScopeId}
            dashboardRefreshKey={dashboardRefreshKey}
            dashboardListSelection={dashboardListSelection}
            onEditCards={(blockId) => {
              setEditingDashboardBlockId(blockId);
              setIsEditingCards(true);
            }}
            onEditSeries={(blockId) => {
              setEditingDashboardBlockId(blockId);
              setIsEditingSeries(true);
            }}
            onSwitchToCharts={handleSwitchDashboardToCharts}
            onCardClick={handleDashboardCardClick}
            onSeriesClick={handleDashboardSeriesClick}
            dashboardAppliesListFilter={dashboardAppliesListFilter}
            onConfigureListContentBlock={
              config.listPageLayout !== undefined
                ? (blockId) => setListPageContentBlockId(blockId)
                : undefined
            }
          />
        )}
      </Stack>

      <CardEditorPanel
        isOpen={isEditingCards}
        listTitle={config.dataSource.title}
        cards={activeListDashboard.cards}
        cardsCount={activeListDashboard.cardsCount}
        dashboardType={activeListDashboard.dashboardType}
        chartType={activeListDashboard.chartType}
        onSave={handleSaveCards}
        onDismiss={() => {
          setIsEditingCards(false);
          setEditingDashboardBlockId(null);
        }}
      />

      <ChartSeriesEditorPanel
        isOpen={isEditingSeries}
        listTitle={config.dataSource.title}
        series={activeListDashboard.chartSeries ?? []}
        dashboardType={activeListDashboard.dashboardType}
        chartType={activeListDashboard.chartType}
        onSave={handleSaveSeries}
        onDismiss={() => {
          setIsEditingSeries(false);
          setEditingDashboardBlockId(null);
        }}
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

      <ListPageLayoutEditorPanel
        isOpen={isEditingPageLayout}
        value={config.listPageLayout ?? defaultListPageLayoutFromLegacy(config)}
        rootDashboard={config.dashboard}
        onSave={handleSaveListPageLayout}
        onDismiss={() => setIsEditingPageLayout(false)}
      />

      <ListPageBlockConfigPanel
        isOpen={listContentBlockPanelOpen}
        block={listContentBlockPanelOpen ? editingListContentBlock : null}
        onDismiss={() => setListPageContentBlockId(null)}
        onApply={handleApplyListContentBlock}
      />
    </>
  );
};

export default DinamicApp;
