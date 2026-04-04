import type {
  IDashboardCardConfig,
  IDashboardConfig,
  IDynamicViewConfig,
  IChartSeriesConfig,
  IListPageBlock,
  IListPageLayoutConfig,
  IListPageSection,
  TChartType,
  TListPageSectionLayout,
} from '../config/types';
import {
  sanitizeBannerConfig,
  sanitizeRichEditorConfig,
} from './listPageBlockConfigUtils';

export const LEGACY_LIST_PAGE_DASHBOARD_BLOCK_ID = 'legacy_dashboard';
export const LEGACY_LIST_PAGE_LIST_BLOCK_ID = 'legacy_list';

export function columnCountForLayout(layout: TListPageSectionLayout): number {
  if (layout === 'one') return 1;
  if (layout === 'two' || layout === 'oneThirdLeft' || layout === 'oneThirdRight') return 2;
  return 3;
}

export function reshapeSectionColumns(section: IListPageSection, newLayout: TListPageSectionLayout): IListPageSection {
  const nc = columnCountForLayout(newLayout);
  const cols = section.columns.map((c) => c.slice());
  while (cols.length < nc) cols.push([]);
  while (cols.length > nc) {
    const tail = cols.pop();
    if (tail && tail.length && cols[0]) cols[0] = cols[0].concat(tail);
    else if (tail && tail.length && !cols[0]) cols[0] = tail.slice();
  }
  return { ...section, layout: newLayout, columns: cols };
}

function newBlockId(): string {
  return `blk_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
}

const VALID_AGG = new Set(['count', 'sum']);

function sanitizeBlockDashboardCard(raw: unknown): IDashboardCardConfig | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const c = raw as Record<string, unknown>;
  if (typeof c.id !== 'string' || typeof c.title !== 'string') return undefined;
  if (typeof c.aggregate !== 'string' || !VALID_AGG.has(c.aggregate)) return undefined;
  if (c.aggregate === 'sum' && (!c.field || typeof c.field !== 'string' || c.field.trim() === '')) return undefined;
  return c as unknown as IDashboardCardConfig;
}

const VALID_CHART: TChartType[] = ['bar', 'line', 'area', 'pie', 'donut'];

function sanitizeBlockDashboard(raw: unknown): IDashboardConfig | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const d = raw as Record<string, unknown>;
  if (typeof d.enabled !== 'boolean') return undefined;
  const dtype = d.dashboardType === 'charts' ? 'charts' : 'cards';
  const cardsCount = typeof d.cardsCount === 'number' && d.cardsCount >= 0 ? d.cardsCount : 0;
  const cardsIn = Array.isArray(d.cards) ? d.cards : [];
  const cards: IDashboardCardConfig[] = [];
  for (let i = 0; i < cardsIn.length; i++) {
    const sc = sanitizeBlockDashboardCard(cardsIn[i]);
    if (sc) cards.push(sc);
  }
  if (dtype === 'cards' && cards.length === 0 && cardsCount === 0) return undefined;
  let chartType: TChartType = 'bar';
  if (typeof d.chartType === 'string' && VALID_CHART.indexOf(d.chartType as TChartType) !== -1) {
    chartType = d.chartType as TChartType;
  }
  let chartSeries: IChartSeriesConfig[] | undefined;
  if (dtype === 'charts' && Array.isArray(d.chartSeries)) {
    const ser: IChartSeriesConfig[] = [];
    for (let j = 0; j < d.chartSeries.length; j++) {
      const e = d.chartSeries[j];
      if (!e || typeof e !== 'object') continue;
      const s = e as Record<string, unknown>;
      if (typeof s.id !== 'string' || typeof s.label !== 'string') continue;
      if (typeof s.aggregate !== 'string' || !VALID_AGG.has(s.aggregate)) continue;
      if (s.aggregate === 'sum' && (!s.field || typeof s.field !== 'string' || (s.field as string).trim() === ''))
        continue;
      ser.push(s as unknown as IChartSeriesConfig);
    }
    chartSeries = ser;
  }
  if (dtype === 'charts') {
    return {
      enabled: d.enabled,
      dashboardType: 'charts',
      cardsCount,
      cards,
      chartType,
      chartSeries: chartSeries ?? [],
    };
  }
  return {
    enabled: d.enabled,
    dashboardType: 'cards',
    cardsCount,
    cards,
    chartType,
  };
}

export function cloneDashboardConfig(d: IDashboardConfig): IDashboardConfig {
  return {
    ...d,
    cards: d.cards.map((c) => ({ ...c, ...(c.style ? { style: { ...c.style } } : {}) })),
    ...(d.chartSeries !== undefined
      ? { chartSeries: d.chartSeries.map((s) => ({ ...s })) }
      : {}),
  };
}

export function countDashboardBlocksInSections(sections: IListPageSection[]): number {
  let n = 0;
  for (let si = 0; si < sections.length; si++) {
    const cols = sections[si].columns;
    for (let ci = 0; ci < cols.length; ci++) {
      const col = cols[ci];
      for (let bi = 0; bi < col.length; bi++) {
        if (col[bi].type === 'dashboard') n += 1;
      }
    }
  }
  return n;
}

export function normalizeListPageLayoutDashboards(
  layout: IListPageLayoutConfig,
  rootDashboard: IDashboardConfig
): IListPageLayoutConfig {
  const n = countDashboardBlocksInSections(layout.sections);
  if (n < 2) return layout;
  return {
    sections: layout.sections.map((sec) => ({
      ...sec,
      columns: sec.columns.map((col) =>
        col.map((b) => {
          if (b.type !== 'dashboard') return b;
          if (b.dashboard !== undefined) return b;
          return { ...b, dashboard: cloneDashboardConfig(rootDashboard) };
        })
      ),
    })),
  };
}

export function findDashboardBlock(
  layout: IListPageLayoutConfig | undefined,
  blockId: string
): IListPageBlock | undefined {
  if (!layout) return undefined;
  for (let si = 0; si < layout.sections.length; si++) {
    const cols = layout.sections[si].columns;
    for (let ci = 0; ci < cols.length; ci++) {
      const col = cols[ci];
      for (let bi = 0; bi < col.length; bi++) {
        const b = col[bi];
        if (b.id === blockId && b.type === 'dashboard') return b;
      }
    }
  }
  return undefined;
}

export function resolveDashboardForListBlock(
  block: IListPageBlock,
  rootDashboard: IDashboardConfig
): IDashboardConfig {
  if (block.type !== 'dashboard') return rootDashboard;
  return block.dashboard ?? rootDashboard;
}

export function getDashboardForEditor(
  config: IDynamicViewConfig,
  blockId: string | null
): IDashboardConfig {
  if (!blockId || !config.listPageLayout) return config.dashboard;
  const found = findDashboardBlock(config.listPageLayout, blockId);
  if (found?.type === 'dashboard') return resolveDashboardForListBlock(found, config.dashboard);
  return config.dashboard;
}

export function findListPageBlockById(
  layout: IListPageLayoutConfig | undefined,
  blockId: string
): IListPageBlock | null {
  if (!layout) return null;
  for (let si = 0; si < layout.sections.length; si++) {
    const cols = layout.sections[si].columns;
    for (let ci = 0; ci < cols.length; ci++) {
      const col = cols[ci];
      for (let bi = 0; bi < col.length; bi++) {
        if (col[bi].id === blockId) return col[bi];
      }
    }
  }
  return null;
}

export function replaceBlockInListPageLayout(
  layout: IListPageLayoutConfig,
  blockId: string,
  next: IListPageBlock
): IListPageLayoutConfig {
  return {
    sections: layout.sections.map((sec) => ({
      ...sec,
      columns: sec.columns.map((col) => col.map((b) => (b.id === blockId ? next : b))),
    })),
  };
}

export function updateDashboardBlockInLayout(
  layout: IListPageLayoutConfig,
  blockId: string,
  next: IDashboardConfig
): IListPageLayoutConfig {
  return {
    sections: layout.sections.map((sec) => ({
      ...sec,
      columns: sec.columns.map((col) =>
        col.map((b) => (b.id === blockId && b.type === 'dashboard' ? { ...b, dashboard: next } : b))
      ),
    })),
  };
}

export function saveDashboardForListBlock(
  config: IDynamicViewConfig,
  blockId: string,
  next: IDashboardConfig
): IDynamicViewConfig {
  const layout = config.listPageLayout;
  if (!layout) {
    return { ...config, dashboard: next };
  }
  const found = findDashboardBlock(layout, blockId);
  if (!found || found.type !== 'dashboard') {
    return { ...config, dashboard: next };
  }
  const n = countDashboardBlocksInSections(layout.sections);
  if (n >= 2 || found.dashboard !== undefined) {
    return {
      ...config,
      listPageLayout: updateDashboardBlockInLayout(layout, blockId, next),
    };
  }
  return { ...config, dashboard: next };
}

export function buildLegacyListPageSections(config: IDynamicViewConfig): IListPageSection[] {
  const showDashboard =
    config.dashboard.enabled &&
    (config.dashboard.dashboardType === 'charts' || config.dashboard.cardsCount > 0);
  const sections: IListPageSection[] = [];
  if (showDashboard) {
    sections.push({
      id: `legacy_${newBlockId()}`,
      layout: 'one',
      columns: [[{ id: LEGACY_LIST_PAGE_DASHBOARD_BLOCK_ID, type: 'dashboard' }]],
    });
  }
  sections.push({
    id: `legacy_${newBlockId()}`,
    layout: 'one',
    columns: [[{ id: LEGACY_LIST_PAGE_LIST_BLOCK_ID, type: 'list' }]],
  });
  return sections;
}

export function getEffectiveListPageSections(config: IDynamicViewConfig): IListPageSection[] {
  const raw = config.listPageLayout?.sections;
  if (raw && raw.length > 0) return raw;
  return buildLegacyListPageSections(config);
}

const VALID_LAYOUTS = new Set<string>(['one', 'two', 'three', 'oneThirdLeft', 'oneThirdRight']);
const VALID_BLOCK_TYPES = new Set<string>(['dashboard', 'list', 'banner', 'editor']);

export function sanitizeListPageLayout(raw: unknown): IListPageLayoutConfig | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const r = raw as Record<string, unknown>;
  if (!Array.isArray(r.sections)) return undefined;
  const sections: IListPageSection[] = [];
  for (let si = 0; si < r.sections.length; si++) {
    const se = r.sections[si];
    if (!se || typeof se !== 'object') continue;
    const s = se as Record<string, unknown>;
    const id = typeof s.id === 'string' && s.id.trim() ? s.id.trim() : `sec_${si}_${Date.now()}`;
    const layoutKey = String(s.layout ?? 'one');
    const layout = (VALID_LAYOUTS.has(layoutKey) ? layoutKey : 'one') as TListPageSectionLayout;
    const need = columnCountForLayout(layout);
    const columnsRaw = Array.isArray(s.columns) ? s.columns : [];
    const columns: IListPageBlock[][] = [];
    for (let ci = 0; ci < need; ci++) {
      const colSrc = columnsRaw[ci];
      const blocks: IListPageBlock[] = [];
      if (Array.isArray(colSrc)) {
        for (let bi = 0; bi < colSrc.length; bi++) {
          const b = colSrc[bi];
          if (!b || typeof b !== 'object') continue;
          const bb = b as Record<string, unknown>;
          const bid = typeof bb.id === 'string' && bb.id.trim() ? bb.id.trim() : newBlockId();
          const bt = String(bb.type ?? '');
          if (!VALID_BLOCK_TYPES.has(bt)) continue;
          const type = bt as IListPageBlock['type'];
          const nestedDash = type === 'dashboard' ? sanitizeBlockDashboard(bb.dashboard) : undefined;
          if (type === 'banner') {
            blocks.push({
              id: bid,
              type,
              banner: sanitizeBannerConfig(bb.banner),
            });
            continue;
          }
          if (type === 'editor') {
            blocks.push({
              id: bid,
              type,
              editor: sanitizeRichEditorConfig(bb.editor),
            });
            continue;
          }
          blocks.push({
            id: bid,
            type,
            ...(nestedDash ? { dashboard: nestedDash } : {}),
          });
        }
      }
      columns.push(blocks);
    }
    sections.push({ id, layout, columns });
  }
  return sections.length > 0 ? { sections } : undefined;
}

export function defaultListPageLayoutFromLegacy(config: IDynamicViewConfig): IListPageLayoutConfig {
  return { sections: buildLegacyListPageSections(config) };
}
