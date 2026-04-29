import * as React from 'react';
import { ActionButton, Separator, Stack } from '@fluentui/react';
import type {
  IDashboardCardConfig,
  IDashboardConfig,
  IChartSeriesConfig,
  IDynamicViewConfig,
  IListPageBlock,
  IListPageSection,
  IListViewFilterConfig,
  TListPageSectionLayout,
} from '../../core/config/types';
import {
  effectiveConfigForListPageBlock,
  resolveDashboardForListBlock,
  sanitizeListPageContentPadding,
} from '../../core/listPage/listPageLayoutUtils';
import { DashboardView } from '../Dashboard/DashboardView';
import { TableView } from '../DataTable/TableView';
import { ListPageAlertBlock } from './ListPageAlertBlock';
import { ListPageBannerBlock } from './ListPageBannerBlock';
import { ListPageButtonsBlock } from './ListPageButtonsBlock';
import { ListPageRichEditorBlock } from './ListPageRichEditorBlock';
import { ListPageSectionTitleBlock } from './ListPageSectionTitleBlock';

export type TListPageDashboardListSelection = {
  blockId: string;
  kind: 'card' | 'series';
  entityId: string;
  filters: IListViewFilterConfig[];
};

export interface IListPageRendererProps {
  config: IDynamicViewConfig;
  sections: IListPageSection[];
  instanceScopeId: string;
  dashboardRefreshKey: number;
  dashboardListSelection: TListPageDashboardListSelection | null;
  onEditCards?: (blockId: string) => void;
  onEditSeries?: (blockId: string) => void;
  onSwitchToCharts?: (blockId: string) => void;
  onCardClick?: (card: IDashboardCardConfig, blockId: string) => void;
  onSeriesClick?: (series: IChartSeriesConfig, blockId: string) => void;
  dashboardAppliesListFilter: boolean;
  /** Abre o painel de configuração do bloco (banner / editor) na página. */
  onConfigureListContentBlock?: (blockId: string) => void;
  /** Exposto acima de cada bloco de tabela (modo lista). */
  onEditTableColumns?: (blockId: string) => void;
  editTableColumnsLabel?: string;
  /** Padding CSS da área do layout (ex. 16px 24px), já validado ao gravar. */
  contentPadding?: string;
  /** Site da página onde o web part está (permissões de modo de lista). */
  pageWebServerRelativeUrl?: string;
  /** Modo de visualização ativo por bloco lista (OData combinado com dashboard vinculado). */
  activeViewModeByBlockId?: Record<string, string>;
  onListViewModeChange?: (listBlockId: string, viewModeId: string) => void;
  onDashboardLinkedTableChange?: (dashboardBlockId: string, pairedListBlockId: string | undefined) => void;
  /** Limpa filtros do dashboard + tabela (passado para o botão "Remover Filtros"). */
  onClearAllFilters?: () => void;
  /** Sinal para resetar filtros internos da TableView. */
  clearTableFiltersSignal?: number;
}

function columnFlexBasis(layout: TListPageSectionLayout, colIndex: number): string {
  if (layout === 'one') return '100%';
  if (layout === 'two') return 'calc(50% - 8px)';
  if (layout === 'three') return 'calc(33.333% - 11px)';
  if (layout === 'oneThirdLeft') return colIndex === 0 ? 'calc(33.333% - 8px)' : 'calc(66.666% - 8px)';
  if (layout === 'oneThirdRight') return colIndex === 0 ? 'calc(66.666% - 8px)' : 'calc(33.333% - 8px)';
  return '100%';
}

function showDashboardBlock(dashboard: IDashboardConfig): boolean {
  return (
    dashboard.enabled &&
    (dashboard.dashboardType === 'charts' || dashboard.cardsCount > 0)
  );
}

function findBlockInSections(sections: IListPageSection[], blockId: string): IListPageBlock | null {
  for (let si = 0; si < sections.length; si++) {
    const cols = sections[si].columns;
    for (let ci = 0; ci < cols.length; ci++) {
      const col = cols[ci];
      for (let bi = 0; bi < col.length; bi++) {
        if (col[bi].id === blockId) return col[bi];
      }
    }
  }
  return null;
}

function tableBlockReceivesDashboardFilters(
  config: IDynamicViewConfig,
  sections: IListPageSection[],
  tableBlock: IListPageBlock,
  selection: TListPageDashboardListSelection | null
): boolean {
  if (!selection?.filters?.length) return false;
  const dashBlock = findBlockInSections(sections, selection.blockId);
  if (!dashBlock || dashBlock.type !== 'dashboard') return false;
  const tTitle = effectiveConfigForListPageBlock(config, tableBlock).dataSource.title;
  const dTitle = effectiveConfigForListPageBlock(config, dashBlock).dataSource.title;
  return tTitle === dTitle;
}

export const ListPageRenderer: React.FC<IListPageRendererProps> = ({
  config,
  sections,
  instanceScopeId,
  dashboardRefreshKey,
  dashboardListSelection,
  onEditCards,
  onEditSeries,
  onSwitchToCharts,
  onCardClick,
  onSeriesClick,
  dashboardAppliesListFilter,
  onConfigureListContentBlock,
  onEditTableColumns,
  editTableColumnsLabel = 'Editar colunas',
  contentPadding,
  pageWebServerRelativeUrl,
  activeViewModeByBlockId = {},
  onListViewModeChange,
  onDashboardLinkedTableChange,
  onClearAllFilters,
  clearTableFiltersSignal,
}) => {
  const rootDash = config.dashboard;
  const layoutPadding = React.useMemo(
    () => sanitizeListPageContentPadding(contentPadding ?? ''),
    [contentPadding]
  );

  const renderBlock = (block: IListPageBlock): React.ReactNode => {
    if (block.type === 'dashboard') {
      const dashCfg = resolveDashboardForListBlock(block, rootDash);
      if (!showDashboardBlock(dashCfg)) return null;
      const sel = dashboardListSelection;
      const selectedCardId =
        sel?.kind === 'card' && sel.blockId === block.id ? sel.entityId : null;
      const selectedSeriesId =
        sel?.kind === 'series' && sel.blockId === block.id ? sel.entityId : null;
      const eff = effectiveConfigForListPageBlock(config, block);
      return (
        <DashboardView
          dashboardBlockId={block.id}
          config={dashCfg}
          dataSource={eff.dataSource}
          refreshKey={dashboardRefreshKey}
          onEditCards={onEditCards}
          onEditSeries={onEditSeries}
          onSwitchToCharts={dashCfg.dashboardType === 'cards' ? onSwitchToCharts : undefined}
          onCardClick={onCardClick}
          selectedCardId={selectedCardId}
          onSeriesClick={onSeriesClick}
          selectedSeriesId={selectedSeriesId}
          dashboardAppliesListFilter={dashboardAppliesListFilter}
          onClearFilters={onClearAllFilters}
          listPairing={{
            rootConfig: config,
            sections,
            dashboardBlock: block,
            rootDashboard: rootDash,
            activeViewModeByBlockId,
          }}
          onLinkedTableChange={
            onDashboardLinkedTableChange !== undefined
              ? (paired) => onDashboardLinkedTableChange(block.id, paired)
              : undefined
          }
        />
      );
    }
    if (block.type === 'list') {
      const effList = effectiveConfigForListPageBlock(config, block);
      const dashFilters =
        tableBlockReceivesDashboardFilters(config, sections, block, dashboardListSelection)
          ? dashboardListSelection?.filters
          : undefined;
      return (
        <Stack key={block.id} tokens={{ childrenGap: 8 }}>
          {onEditTableColumns !== undefined ? (
            <Stack horizontal horizontalAlign="end" styles={{ root: { width: '100%' } }}>
              <ActionButton
                iconProps={{ iconName: 'ColumnOptions' }}
                onClick={() => onEditTableColumns(block.id)}
                styles={{ root: { height: 30, color: '#0078d4' } }}
              >
                {editTableColumnsLabel}
              </ActionButton>
            </Stack>
          ) : null}
          <TableView
            config={effList}
            dashboardListFilters={dashFilters}
            instanceScopeId={`${instanceScopeId}_${block.id}`}
            pageWebServerRelativeUrl={pageWebServerRelativeUrl}
            onActiveViewModeChange={
              onListViewModeChange !== undefined
                ? (modeId) => onListViewModeChange(block.id, modeId)
                : undefined
            }
            clearFiltersSignal={clearTableFiltersSignal}
          />
        </Stack>
      );
    }
    if (block.type === 'banner') {
      return (
        <ListPageBannerBlock
          key={block.id}
          banner={block.banner}
          onConfigure={
            onConfigureListContentBlock !== undefined
              ? () => onConfigureListContentBlock(block.id)
              : undefined
          }
        />
      );
    }
    if (block.type === 'editor') {
      return (
        <ListPageRichEditorBlock
          key={block.id}
          editor={block.editor}
          onConfigure={
            onConfigureListContentBlock !== undefined
              ? () => onConfigureListContentBlock(block.id)
              : undefined
          }
        />
      );
    }
    if (block.type === 'sectionTitle') {
      return (
        <ListPageSectionTitleBlock
          key={block.id}
          sectionTitle={block.sectionTitle}
          onConfigure={
            onConfigureListContentBlock !== undefined
              ? () => onConfigureListContentBlock(block.id)
              : undefined
          }
        />
      );
    }
    if (block.type === 'alert') {
      const effAlert = effectiveConfigForListPageBlock(config, block);
      return (
        <ListPageAlertBlock
          key={block.id}
          alert={block.alert}
          listTitle={effAlert.dataSource.title ?? ''}
          onConfigure={
            onConfigureListContentBlock !== undefined
              ? () => onConfigureListContentBlock(block.id)
              : undefined
          }
        />
      );
    }
    if (block.type === 'buttons') {
      return (
        <ListPageButtonsBlock
          key={block.id}
          buttons={block.buttons}
          onConfigure={
            onConfigureListContentBlock !== undefined
              ? () => onConfigureListContentBlock(block.id)
              : undefined
          }
        />
      );
    }
    return null;
  };

  const inner = (
    <>
      {sections.map((section, si) => (
        <React.Fragment key={section.id}>
          {si > 0 ? <Separator styles={{ root: { marginTop: 8, marginBottom: 16 } }} /> : null}
          <div
            style={{
              display: 'flex',
              flexDirection: 'row',
              flexWrap: 'wrap',
              gap: 16,
              width: '100%',
              alignItems: 'flex-start',
            }}
          >
            {section.columns.map((blocks, ci) => (
              <div
                key={`${section.id}_c${ci}`}
                style={{
                  flex: `1 1 ${columnFlexBasis(section.layout, ci)}`,
                  minWidth: section.layout === 'three' ? 200 : 220,
                  maxWidth: '100%',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: 16,
                }}
              >
                {blocks.map((b) => (
                  <div key={b.id}>{renderBlock(b)}</div>
                ))}
              </div>
            ))}
          </div>
        </React.Fragment>
      ))}
    </>
  );

  const blocksScopeClass = `dinamicSxLayoutScope_${instanceScopeId.replace(/[^a-zA-Z0-9_-]/g, '_')}`;
  const rawBlocksCss = config.listPageLayout?.customBlocksCss ?? '';
  const scopedBlocksCss = rawBlocksCss.trim()
    ? rawBlocksCss.replace(/\.dinamicSx/g, `.${blocksScopeClass} .dinamicSx`)
    : '';
  const blocksCssTag = scopedBlocksCss
    ? <style type="text/css">{scopedBlocksCss}</style>
    : null;

  if (layoutPadding) {
    return (
      <div className={blocksScopeClass} style={{ padding: layoutPadding, width: '100%', maxWidth: '100%', boxSizing: 'border-box' }}>
        {blocksCssTag}
        {inner}
      </div>
    );
  }

  return (
    <div className={blocksScopeClass} style={{ width: '100%' }}>
      {blocksCssTag}
      {inner}
    </div>
  );
};
