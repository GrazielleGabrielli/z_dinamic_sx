import * as React from 'react';
import { useState, useCallback, useEffect, useMemo } from 'react';
import { Text, Stack, ActionButton } from '@fluentui/react';
import type { IDinamicAppProps } from './IDinamicAppProps';
import { coerceDashboardShape, parseConfig } from '../core/config/validators';
import {
  IDashboardCardConfig,
  IDashboardConfig,
  IChartSeriesConfig,
  IDynamicViewConfig,
  IFormManagerConfig,
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
  findListPageBlockInSections,
  getDashboardForEditor,
  getEffectiveListPageSections,
  mergeLinkedListMemoryIntoConfig,
  LEGACY_LIST_PAGE_DASHBOARD_BLOCK_ID,
  replaceBlockInListPageLayout,
  saveDashboardForListBlock,
} from '../core/listPage/listPageLayoutUtils';
import { upsertConfigMemoryForListSource } from '../core/config/configMemory';
import { effectiveDashboardFilters } from '../core/dashboard/effectiveDashboardFilters';
import {
  chartSeriesToDashboardCards,
  dashboardCardsToChartSeries,
} from '../core/dashboard/chartSeriesToDashboardCards';
import { generateDefaultCards, getDefaultFormManagerConfig } from '../core/config/utils';
import { ConfigWizard } from './Wizard/ConfigWizard';
import { CardEditorPanel } from './Dashboard/CardEditor/CardEditorPanel';
import { ChartSeriesEditorPanel } from './Dashboard/ChartEditor/ChartSeriesEditorPanel';
import { TableColumnsEditorPanel } from './DataTable/TableColumnsEditorPanel';
import { ProjectManagementView } from './ProjectManagement/ProjectManagementView';
import { ListPageRenderer, type TListPageDashboardListSelection } from './ListPage/ListPageRenderer';
import { ListPageLayoutEditorPanel } from './ListPage/ListPageLayoutEditorPanel';
import { ListPageBlockConfigPanel } from './ListPage/ListPageBlockConfigPanel';
import { FormManagerView } from './FormManager/FormManagerView';
import { FormManagerConfigPanel } from './FormManager/FormManagerConfigPanel';
import { PersistStatusBar } from './PersistStatusBar';

const DinamicApp: React.FC<IDinamicAppProps> = ({
  configJson,
  siteUrl,
  instanceScopeId,
  onSaveConfig,
  persistStatus,
  forcedMode,
}) => {
  const [isEditingWebPart, setIsEditingWebPart] = useState(false);
  const [isEditingCards, setIsEditingCards] = useState(false);
  const [isEditingSeries, setIsEditingSeries] = useState(false);
  const [isEditingTableColumns, setIsEditingTableColumns] = useState(false);
  const [isEditingPageLayout, setIsEditingPageLayout] = useState(false);
  const [isEditingFormManager, setIsEditingFormManager] = useState(false);
  const [listPageContentBlockId, setListPageContentBlockId] = useState<string | null>(null);
  const [editingDashboardBlockId, setEditingDashboardBlockId] = useState<string | null>(null);
  const [editingTableBlockId, setEditingTableBlockId] = useState<string | null>(null);
  const [dashboardRefreshKey, setDashboardRefreshKey] = useState(0);
  const [dashboardListSelection, setDashboardListSelection] =
    useState<TListPageDashboardListSelection | null>(null);

  const rawConfig = useMemo(() => parseConfig(configJson ?? undefined), [configJson]);
  const config = useMemo((): IDynamicViewConfig | undefined => {
    if (forcedMode === undefined) return rawConfig;
    if (rawConfig === undefined) return undefined;
    return { ...rawConfig, mode: forcedMode };
  }, [rawConfig, forcedMode]);

  // 'pending' não bloqueia — em Edit Mode o usuário pode alterar a config múltiplas vezes antes de salvar a página
  const isSaving = persistStatus === 'saving' || persistStatus === 'persisting';

  const saveConfig = useCallback(
    (cfg: IDynamicViewConfig) => {
      if (isSaving) {
        console.warn('[DinamicSX Persist] save bloqueado no componente — persistência em andamento');
        return;
      }
      const next: IDynamicViewConfig =
        forcedMode !== undefined ? { ...cfg, mode: forcedMode } : cfg;
      onSaveConfig(next);
    },
    [forcedMode, isSaving, onSaveConfig]
  );

  const formManagerResolved = useMemo(
    () => config?.formManager ?? getDefaultFormManagerConfig(),
    [config?.formManager]
  );

  const tableEditorSource = useMemo(() => {
    if (!config) {
      return {
        listTitle: '',
        listView: undefined as IListViewConfig | undefined,
        pagination: undefined as IPaginationConfig | undefined,
        pdfTemplate: undefined as import('../core/config/types').IPdfTemplateConfig | undefined,
      };
    }
    const bid = editingTableBlockId;
    if (!bid) {
      return {
        listTitle: config.dataSource.title,
        listView: config.listView,
        pagination: config.pagination,
        pdfTemplate: config.pdfTemplate,
      };
    }
    const block = findListPageBlockInSections(getEffectiveListPageSections(config), bid);
    const lt = block?.linkedListBinding?.listTitle?.trim();
    if (!lt) {
      return {
        listTitle: config.dataSource.title,
        listView: config.listView,
        pagination: config.pagination,
        pdfTemplate: config.pdfTemplate,
      };
    }
    const eff = mergeLinkedListMemoryIntoConfig(config, { kind: 'list', title: lt });
    return {
      listTitle: eff.dataSource.title,
      listView: eff.listView,
      pagination: eff.pagination,
      pdfTemplate: eff.pdfTemplate,
    };
  }, [config, editingTableBlockId]);

  const dashboardEditorListTitle = useMemo(() => {
    if (!config) return '';
    const bid = editingDashboardBlockId;
    if (!bid) return config.dataSource.title;
    const b = findListPageBlockInSections(getEffectiveListPageSections(config), bid);
    const lt = b?.linkedListBinding?.listTitle?.trim();
    return lt || config.dataSource.title;
  }, [config, editingDashboardBlockId]);

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
    saveConfig(newConfig);
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
    saveConfig(saveDashboardForListBlock(config, blockId, coerceDashboardShape(next)));
  }, [config, saveConfig]);

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
    saveConfig(saveDashboardForListBlock(config, bid, coerceDashboardShape(dashboard)));
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
    saveConfig(saveDashboardForListBlock(config, bid, coerceDashboardShape(dashboard)));
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
    const bid = editingTableBlockId;
    if (bid) {
      const block = findListPageBlockInSections(getEffectiveListPageSections(config), bid);
      const lt = block?.linkedListBinding?.listTitle?.trim();
      if (lt) {
        saveConfig(
          upsertConfigMemoryForListSource(
            config,
            { kind: 'list', title: lt },
            {
              listView,
              pagination,
              ...(pdfTemplate !== undefined ? { pdfTemplate } : {}),
            }
          )
        );
        setIsEditingTableColumns(false);
        setEditingTableBlockId(null);
        return;
      }
    }
    saveConfig({
      ...config,
      listView,
      pagination,
      ...(projectManagement !== undefined && { projectManagement }),
      ...(pdfTemplate !== undefined && { pdfTemplate }),
    });
    setIsEditingTableColumns(false);
    setEditingTableBlockId(null);
  };

  const handleSaveListPageLayout = (layout: IListPageLayoutConfig): void => {
    if (!config) return;
    saveConfig({
      ...config,
      listPageLayout: layout,
    });
    setIsEditingPageLayout(false);
  };

  const handleSaveFormManagerConfig = (formManager: IFormManagerConfig): void => {
    if (!config) return;
    saveConfig({ ...config, formManager });
    setIsEditingFormManager(false);
  };

  const handleApplyListContentBlock = (next: IListPageBlock): void => {
    if (!config?.listPageLayout) return;
    saveConfig({
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
      editingListContentBlock.type === 'alert' ||
      editingListContentBlock.type === 'buttons');

  if (config === undefined || isEditingWebPart) {
    return (
      <ConfigWizard
        siteUrl={siteUrl}
        onComplete={handleWizardComplete}
        initialValues={config}
        onCancel={config !== undefined ? () => setIsEditingWebPart(false) : undefined}
        forcedMode={forcedMode}
      />
    );
  }

  const activeListDashboard = getDashboardForEditor(config, editingDashboardBlockId);

  return (
    <>
      <PersistStatusBar status={persistStatus} />

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
          disabled={isSaving}
          styles={{ root: { color: '#605e5c', fontSize: 12 } }}
        >
          Editar configuração
        </ActionButton>
      </div>

      <Stack styles={{ root: { padding: '20px 24px 0' } }}>
        <Stack
          horizontal
          horizontalAlign={config.mode === 'formManager' ? 'end' : 'space-between'}
          verticalAlign="center"
          tokens={{ childrenGap: 8 }}
          styles={{ root: { padding: '16px 0 8px' } }}
        >
          {config.mode !== 'formManager' && (
            <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
              {config.dataSource.title}
            </Text>
          )}
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
            {config.mode === 'formManager' && (
              <ActionButton
                iconProps={{ iconName: 'FormLibrary' }}
                onClick={() => setIsEditingFormManager(true)}
                styles={{ root: { height: 28, color: '#0078d4' } }}
              >
                Configurar formulário
              </ActionButton>
            )}
          </Stack>
        </Stack>

        {config.mode === 'formManager' ? (
          <FormManagerView config={config} />
        ) : config.mode === 'projectManagement' ? (
          <ProjectManagementView
            config={config}
            dashboardListFilters={dashboardListSelection?.filters}
            onItemUpdated={triggerDashboardRefresh}
            onEditTableColumns={() => setIsEditingTableColumns(true)}
          />
        ) : (
          <ListPageRenderer
            config={config}
            sections={getEffectiveListPageSections(config)}
            instanceScopeId={instanceScopeId}
            dashboardRefreshKey={dashboardRefreshKey}
            dashboardListSelection={dashboardListSelection}
            contentPadding={config.listPageLayout?.contentPadding}
            onEditTableColumns={(blockId) => {
              setEditingTableBlockId(blockId);
              setIsEditingTableColumns(true);
            }}
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
        listTitle={dashboardEditorListTitle}
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
        listTitle={dashboardEditorListTitle}
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
        listTitle={tableEditorSource.listTitle}
        listView={tableEditorSource.listView ?? config.listView}
        pagination={tableEditorSource.pagination ?? config.pagination}
        projectManagement={config.projectManagement}
        pdfTemplate={tableEditorSource.pdfTemplate ?? config.pdfTemplate}
        onSave={handleSaveTableColumns}
        onDismiss={() => {
          setIsEditingTableColumns(false);
          setEditingTableBlockId(null);
        }}
      />

      <ListPageLayoutEditorPanel
        isOpen={isEditingPageLayout}
        value={config.listPageLayout ?? defaultListPageLayoutFromLegacy(config)}
        rootDashboard={config.dashboard}
        sourceListTitle={config.dataSource.title ?? ''}
        onSave={handleSaveListPageLayout}
        onDismiss={() => setIsEditingPageLayout(false)}
      />

      <ListPageBlockConfigPanel
        isOpen={listContentBlockPanelOpen}
        block={listContentBlockPanelOpen ? editingListContentBlock : null}
        listTitle={
          editingListContentBlock?.type === 'alert' &&
          editingListContentBlock.linkedListBinding?.listTitle?.trim()
            ? editingListContentBlock.linkedListBinding.listTitle.trim()
            : config.dataSource.title ?? ''
        }
        onDismiss={() => setListPageContentBlockId(null)}
        onApply={handleApplyListContentBlock}
      />

      {config.mode === 'formManager' && (
        <FormManagerConfigPanel
          isOpen={isEditingFormManager}
          listTitle={config.dataSource.title}
          value={formManagerResolved}
          onSave={handleSaveFormManagerConfig}
          onDismiss={() => setIsEditingFormManager(false)}
        />
      )}
    </>
  );
};

export default DinamicApp;
