import * as React from 'react';
import { Separator } from '@fluentui/react';
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
import { resolveDashboardForListBlock } from '../../core/listPage/listPageLayoutUtils';
import { DashboardView } from '../Dashboard/DashboardView';
import { TableView } from '../DataTable/TableView';
import { ListPageAlertBlock } from './ListPageAlertBlock';
import { ListPageBannerBlock } from './ListPageBannerBlock';
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
  onEditCards: (blockId: string) => void;
  onEditSeries: (blockId: string) => void;
  onSwitchToCharts?: (blockId: string) => void;
  onCardClick?: (card: IDashboardCardConfig, blockId: string) => void;
  onSeriesClick?: (series: IChartSeriesConfig, blockId: string) => void;
  dashboardAppliesListFilter: boolean;
  /** Abre o painel de configuração do bloco (banner / editor) na página. */
  onConfigureListContentBlock?: (blockId: string) => void;
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
}) => {
  const rootDash = config.dashboard;

  const renderBlock = (block: IListPageBlock): React.ReactNode => {
    if (block.type === 'dashboard') {
      const dashCfg = resolveDashboardForListBlock(block, rootDash);
      if (!showDashboardBlock(dashCfg)) return null;
      const sel = dashboardListSelection;
      const selectedCardId =
        sel?.kind === 'card' && sel.blockId === block.id ? sel.entityId : null;
      const selectedSeriesId =
        sel?.kind === 'series' && sel.blockId === block.id ? sel.entityId : null;
      return (
        <DashboardView
          dashboardBlockId={block.id}
          config={dashCfg}
          dataSource={config.dataSource}
          refreshKey={dashboardRefreshKey}
          onEditCards={onEditCards}
          onEditSeries={onEditSeries}
          onSwitchToCharts={dashCfg.dashboardType === 'cards' ? onSwitchToCharts : undefined}
          onCardClick={onCardClick}
          selectedCardId={selectedCardId}
          onSeriesClick={onSeriesClick}
          selectedSeriesId={selectedSeriesId}
          dashboardAppliesListFilter={dashboardAppliesListFilter}
        />
      );
    }
    if (block.type === 'list') {
      return (
        <TableView
          key={block.id}
          config={config}
          dashboardListFilters={dashboardListSelection?.filters}
          instanceScopeId={`${instanceScopeId}_${block.id}`}
        />
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
      return (
        <ListPageAlertBlock
          key={block.id}
          alert={block.alert}
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

  return (
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
};
