import type { IDashboardConfig, IDynamicViewConfig, IListViewFilterConfig } from '../config/types';
import {
  effectiveConfigForListPageBlock,
  findBlockInSections,
  findMatchingListBlockIdForDashboard,
  getEffectiveListPageSections,
} from '../listPage/listPageLayoutUtils';
import { getViewModeFiltersById } from '../listView/buildListQuery';

export function resolveDashboardCombineWithViewMode(dash: IDashboardConfig): boolean {
  if (dash.combineWithActiveViewMode === true) return true;
  if (dash.combineWithActiveViewMode === false) return false;
  if (dash.dashboardType === 'charts') {
    return (dash.chartSeries ?? []).some((s) => s.combineWithActiveViewMode === true);
  }
  return (dash.cards ?? []).some((c) => c.combineWithActiveViewMode === true);
}

export function mergeDashboardClickFiltersWithViewMode(
  cfg: IDynamicViewConfig,
  combine: boolean | undefined,
  dashboardBlockId: string,
  dashboardFilters: IListViewFilterConfig[],
  listViewModeByBlockId: Record<string, string>
): IListViewFilterConfig[] {
  if (combine !== true) return dashboardFilters;
  const sections = getEffectiveListPageSections(cfg);
  const listBlockId = findMatchingListBlockIdForDashboard(cfg, sections, dashboardBlockId);
  if (!listBlockId) return dashboardFilters;
  const lb = findBlockInSections(sections, listBlockId);
  if (!lb || lb.type !== 'list') return dashboardFilters;
  const eff = effectiveConfigForListPageBlock(cfg, lb);
  const lv = eff.listView;
  const modeId = listViewModeByBlockId[listBlockId] ?? lv.activeViewModeId ?? lv.viewModes?.[0]?.id;
  const vm = getViewModeFiltersById(lv, modeId);
  return [...vm, ...dashboardFilters];
}
