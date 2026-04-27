import type {
  IDashboardConfig,
  IDataSourceConfig,
  IDynamicViewConfig,
  IListPageBlock,
  IListPageSection,
} from '../config/types';
import type { IFieldMetadata } from '../../../../services';
import type { IDynamicContext } from '../dynamicTokens/types';
import { buildListFilter, getActiveViewModeFilters } from '../listView';
import {
  effectiveConfigForListPageBlock,
  findBlockInSections,
} from './listPageLayoutUtils';

function normWeb(u: string | undefined): string {
  return (u ?? '').trim().replace(/\/+$/, '');
}

export function dataSourcesMatch(a: IDataSourceConfig, b: IDataSourceConfig): boolean {
  const ta = (a.title ?? '').trim();
  const tb = (b.title ?? '').trim();
  if (ta !== tb) return false;
  return normWeb(a.webServerRelativeUrl) === normWeb(b.webServerRelativeUrl);
}

export function buildLinkedViewModeOData(
  rootConfig: IDynamicViewConfig,
  sections: IListPageSection[],
  dashboardDataSource: IDataSourceConfig,
  linkedListBlockId: string | undefined,
  activeViewModeByBlockId: Record<string, string>,
  dynamicContext?: IDynamicContext,
  fieldsMetadata?: IFieldMetadata[]
): string | undefined {
  const lid = linkedListBlockId?.trim();
  if (!lid || lid === '') return undefined;
  const block = findBlockInSections(sections, lid);
  if (!block || block.type !== 'list') return undefined;
  const eff = effectiveConfigForListPageBlock(rootConfig, block);
  if (!dataSourcesMatch(eff.dataSource, dashboardDataSource)) return undefined;
  const lv = eff.listView;
  const modes = lv.viewModes ?? [];
  const desired =
    activeViewModeByBlockId[lid] ??
    lv.activeViewModeId ??
    modes[0]?.id ??
    'all';
  const lvWithMode = { ...lv, activeViewModeId: desired };
  const filters = getActiveViewModeFilters(lvWithMode);
  return buildListFilter(filters, { dynamicContext, fieldsMetadata });
}

export function resolveLinkedListBlockIdForDashboard(
  dashboardBlock: IListPageBlock,
  rootDashboard: IDashboardConfig
): string | undefined {
  if (dashboardBlock.type === 'dashboard') {
    const p = dashboardBlock.pairedListBlockId?.trim();
    if (p) return p;
  }
  const l = rootDashboard.linkedListBlockId?.trim();
  return l || undefined;
}

export function collectCompatibleListBlocksForDashboard(
  rootConfig: IDynamicViewConfig,
  sections: IListPageSection[],
  dashboardBlock: IListPageBlock
): { id: string; label: string }[] {
  if (dashboardBlock.type !== 'dashboard') return [];
  const effDash = effectiveConfigForListPageBlock(rootConfig, dashboardBlock);
  const out: { id: string; label: string }[] = [];
  let n = 0;
  for (let si = 0; si < sections.length; si++) {
    const cols = sections[si].columns;
    for (let ci = 0; ci < cols.length; ci++) {
      const col = cols[ci];
      for (let bi = 0; bi < col.length; bi++) {
        const b = col[bi];
        if (b.type !== 'list') continue;
        const eff = effectiveConfigForListPageBlock(rootConfig, b);
        if (!dataSourcesMatch(eff.dataSource, effDash.dataSource)) continue;
        n += 1;
        out.push({ id: b.id, label: `Tabela ${n}` });
      }
    }
  }
  return out;
}
