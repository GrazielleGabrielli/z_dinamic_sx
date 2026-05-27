import type {
  IDataSourceConfig,
  IDashboardCardConfig,
  IDashboardCardLayoutStyle,
  IDashboardConfig,
  IDynamicViewConfig,
  IChartSeriesConfig,
  IListPageBlock,
  IListPageLayoutConfig,
  IListPageLinkedListBinding,
  IListPageSection,
  TChartType,
  TListPageSectionLayout,
} from '../config/types';
import { sourceKey } from '../config/configMemory';
import { getDefaultDashboardCardLayoutStyle } from '../dashboard/utils';
import {
  sanitizeAlertConfig,
  sanitizeBannerConfig,
  sanitizeButtonsConfig,
  sanitizeRichEditorConfig,
  sanitizeSectionTitleConfig,
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

/** Alinha `columns.length` ao `layout`; colunas a mais são fundidas na primeira (igual a `reshapeSectionColumns`). */
export function mergeExtraListPageColumnsIntoLayout(
  layout: TListPageSectionLayout,
  columns: IListPageBlock[][]
): IListPageBlock[][] {
  return reshapeSectionColumns({ id: '_merge', layout, columns }, layout).columns;
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
  const linkedListBlockId =
    typeof d.linkedListBlockId === 'string' && d.linkedListBlockId.trim().length > 0
      ? d.linkedListBlockId.trim()
      : undefined;
  const dashCombine =
    d.combineWithActiveViewMode === true
      ? true
      : d.combineWithActiveViewMode === false
        ? false
        : undefined;
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
  const cardLayoutStyle = sanitizeBlockCardLayoutStyle(d.cardLayoutStyle);
  if (dtype === 'charts') {
    return {
      enabled: d.enabled,
      dashboardType: 'charts',
      cardsCount,
      cards,
      chartType,
      chartSeries: chartSeries ?? [],
      ...(cardLayoutStyle ? { cardLayoutStyle } : {}),
      ...(dashCombine !== undefined && { combineWithActiveViewMode: dashCombine }),
      ...(linkedListBlockId ? { linkedListBlockId } : {}),
    };
  }
  return {
    enabled: d.enabled,
    dashboardType: 'cards',
    cardsCount,
    cards,
    chartType,
    ...(cardLayoutStyle ? { cardLayoutStyle } : {}),
    ...(dashCombine !== undefined && { combineWithActiveViewMode: dashCombine }),
    ...(linkedListBlockId ? { linkedListBlockId } : {}),
  };
}

const VALID_VARIANTS = new Set(['default', 'outlined', 'soft', 'solid']);
const VALID_BORDER_RADIUS = new Set(['none', 'sm', 'md', 'lg', 'xl', 'full']);
const VALID_PADDING = new Set(['sm', 'md', 'lg']);
const VALID_SHADOW = new Set(['none', 'sm', 'md', 'lg']);
const VALID_TITLE_SIZE = new Set(['xs', 'sm', 'md', 'lg']);
const VALID_SUBTITLE_SIZE = new Set(['xs', 'sm', 'md']);
const VALID_VALUE_SIZE = new Set(['lg', 'xl', '2xl', '3xl']);
const VALID_FONT_WEIGHT = new Set(['normal', 'medium', 'semibold', 'bold']);
const VALID_ALIGN = new Set(['left', 'center', 'right']);
const VALID_ICON_POSITION = new Set(['left', 'top', 'right']);
const VALID_LOADING_STYLE = new Set(['skeleton', 'spinner', 'text']);

function sanitizeBlockCardLayoutStyle(raw: unknown): IDashboardCardLayoutStyle | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const s = raw as Record<string, unknown>;
  const variant = VALID_VARIANTS.has(s.variant as string)
    ? (s.variant as IDashboardCardLayoutStyle['variant'])
    : undefined;
  const borderRadius = VALID_BORDER_RADIUS.has(s.borderRadius as string)
    ? (s.borderRadius as IDashboardCardLayoutStyle['borderRadius'])
    : undefined;
  const padding = VALID_PADDING.has(s.padding as string)
    ? (s.padding as IDashboardCardLayoutStyle['padding'])
    : undefined;
  const shadow = VALID_SHADOW.has(s.shadow as string)
    ? (s.shadow as IDashboardCardLayoutStyle['shadow'])
    : undefined;
  const border = typeof s.border === 'boolean' ? s.border : undefined;
  const titleSize = VALID_TITLE_SIZE.has(s.titleSize as string)
    ? (s.titleSize as IDashboardCardLayoutStyle['titleSize'])
    : undefined;
  const subtitleSize = VALID_SUBTITLE_SIZE.has(s.subtitleSize as string)
    ? (s.subtitleSize as IDashboardCardLayoutStyle['subtitleSize'])
    : undefined;
  const valueSize = VALID_VALUE_SIZE.has(s.valueSize as string)
    ? (s.valueSize as IDashboardCardLayoutStyle['valueSize'])
    : undefined;
  const titleWeight = VALID_FONT_WEIGHT.has(s.titleWeight as string)
    ? (s.titleWeight as IDashboardCardLayoutStyle['titleWeight'])
    : undefined;
  const valueWeight = VALID_FONT_WEIGHT.has(s.valueWeight as string)
    ? (s.valueWeight as IDashboardCardLayoutStyle['valueWeight'])
    : undefined;
  const align = VALID_ALIGN.has(s.align as string)
    ? (s.align as IDashboardCardLayoutStyle['align'])
    : undefined;
  const iconPosition = VALID_ICON_POSITION.has(s.iconPosition as string)
    ? (s.iconPosition as IDashboardCardLayoutStyle['iconPosition'])
    : undefined;
  const loadingStyle = VALID_LOADING_STYLE.has(s.loadingStyle as string)
    ? (s.loadingStyle as IDashboardCardLayoutStyle['loadingStyle'])
    : undefined;
  const showSubtitle = typeof s.showSubtitle === 'boolean' ? s.showSubtitle : undefined;
  const showValue = typeof s.showValue === 'boolean' ? s.showValue : undefined;
  const showIcon = typeof s.showIcon === 'boolean' ? s.showIcon : undefined;
  if (
    variant === undefined &&
    borderRadius === undefined &&
    padding === undefined &&
    shadow === undefined &&
    border === undefined &&
    titleSize === undefined &&
    subtitleSize === undefined &&
    valueSize === undefined &&
    titleWeight === undefined &&
    valueWeight === undefined &&
    align === undefined &&
    iconPosition === undefined &&
    loadingStyle === undefined &&
    showSubtitle === undefined &&
    showValue === undefined &&
    showIcon === undefined
  ) {
    return undefined;
  }
  return {
    ...getDefaultDashboardCardLayoutStyle(),
    ...(variant !== undefined ? { variant } : {}),
    ...(borderRadius !== undefined ? { borderRadius } : {}),
    ...(padding !== undefined ? { padding } : {}),
    ...(shadow !== undefined ? { shadow } : {}),
    ...(border !== undefined ? { border } : {}),
    ...(titleSize !== undefined ? { titleSize } : {}),
    ...(subtitleSize !== undefined ? { subtitleSize } : {}),
    ...(valueSize !== undefined ? { valueSize } : {}),
    ...(titleWeight !== undefined ? { titleWeight } : {}),
    ...(valueWeight !== undefined ? { valueWeight } : {}),
    ...(align !== undefined ? { align } : {}),
    ...(iconPosition !== undefined ? { iconPosition } : {}),
    ...(loadingStyle !== undefined ? { loadingStyle } : {}),
    ...(showSubtitle !== undefined ? { showSubtitle } : {}),
    ...(showValue !== undefined ? { showValue } : {}),
    ...(showIcon !== undefined ? { showIcon } : {}),
  };
}

export function cloneDashboardConfig(d: IDashboardConfig): IDashboardConfig {
  return {
    ...d,
    cards: d.cards.map((c) => ({ ...c, ...(c.style ? { style: { ...c.style } } : {}) })),
    ...(d.cardLayoutStyle !== undefined ? { cardLayoutStyle: { ...d.cardLayoutStyle } } : {}),
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
    ...layout,
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
  return findListPageBlockInSections(layout.sections, blockId);
}

export function findListPageBlockInSections(
  sections: IListPageSection[] | undefined,
  blockId: string
): IListPageBlock | null {
  if (!sections) return null;
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

export function replaceBlockInListPageLayout(
  layout: IListPageLayoutConfig,
  blockId: string,
  next: IListPageBlock
): IListPageLayoutConfig {
  return {
    ...layout,
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
    ...layout,
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

export function findBlockInSections(sections: IListPageSection[], blockId: string): IListPageBlock | null {
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

function flattenBlocksReadingOrder(sections: IListPageSection[]): IListPageBlock[] {
  const out: IListPageBlock[] = [];
  for (let si = 0; si < sections.length; si++) {
    const cols = sections[si].columns;
    for (let ci = 0; ci < cols.length; ci++) {
      const col = cols[ci];
      for (let bi = 0; bi < col.length; bi++) {
        out.push(col[bi]);
      }
    }
  }
  return out;
}

/** Bloco `list` na página que usa a mesma lista que o dashboard `dashboardBlockId`. */
export function findMatchingListBlockIdForDashboard(
  config: IDynamicViewConfig,
  sections: IListPageSection[],
  dashboardBlockId: string
): string | undefined {
  const dashBlock = findBlockInSections(sections, dashboardBlockId);
  if (!dashBlock || dashBlock.type !== 'dashboard') return undefined;
  const paired = dashBlock.pairedListBlockId?.trim();
  if (paired) {
    const pb = findBlockInSections(sections, paired);
    if (pb?.type === 'list') return paired;
  }
  const dTitle = effectiveConfigForListPageBlock(config, dashBlock).dataSource.title?.trim();
  if (!dTitle) return undefined;
  const flat = flattenBlocksReadingOrder(sections);
  const dashIdx = flat.findIndex((b) => b.id === dashboardBlockId);
  const matching: { id: string; idx: number }[] = [];
  for (let i = 0; i < flat.length; i++) {
    const b = flat[i];
    if (b.type !== 'list') continue;
    const t = effectiveConfigForListPageBlock(config, b).dataSource.title?.trim();
    if (t === dTitle) matching.push({ id: b.id, idx: i });
  }
  if (matching.length === 0) return undefined;
  if (matching.length === 1) return matching[0].id;
  if (dashIdx >= 0) {
    const afterDash = matching.filter((m) => m.idx > dashIdx);
    if (afterDash.length > 0) return afterDash[0].id;
  }
  return matching[0].id;
}

export function sanitizeListPageContentPadding(raw: unknown): string | undefined {
  if (typeof raw !== 'string') return undefined;
  const t = raw.trim();
  if (!t) return undefined;
  const parts = t.split(/\s+/).filter(Boolean);
  if (parts.length === 0 || parts.length > 4) return undefined;
  for (let i = 0; i < parts.length; i++) {
    if (!/^(0|[1-9]\d{0,3})px$/i.test(parts[i])) return undefined;
  }
  return parts.map((p) => p.toLowerCase()).join(' ');
}

const VALID_LAYOUTS = new Set<string>(['one', 'two', 'three', 'oneThirdLeft', 'oneThirdRight']);
const VALID_BLOCK_TYPES = new Set<string>([
  'dashboard',
  'list',
  'banner',
  'editor',
  'sectionTitle',
  'alert',
  'buttons',
]);

export function sanitizeLinkedListBinding(
  raw: unknown,
  blockType: IListPageBlock['type']
): IListPageLinkedListBinding | undefined {
  if (blockType !== 'dashboard' && blockType !== 'list' && blockType !== 'alert') return undefined;
  if (!raw || typeof raw !== 'object') return undefined;
  const o = raw as Record<string, unknown>;
  const listTitle = typeof o.listTitle === 'string' ? o.listTitle.trim() : '';
  const parentLookupFieldInternalName =
    typeof o.parentLookupFieldInternalName === 'string' ? o.parentLookupFieldInternalName.trim() : '';
  if (!listTitle || !parentLookupFieldInternalName) return undefined;
  return { listTitle, parentLookupFieldInternalName };
}

/** Aplica só listView / pagination / tableConfig da memória da lista filha (não dashboard global nem layout). */
export function mergeLinkedListMemoryIntoConfig(
  base: IDynamicViewConfig,
  childDataSource: IDataSourceConfig
): IDynamicViewConfig {
  const key = sourceKey(childDataSource);
  const snap = base.configMemory?.bySource?.[key]?.list;
  if (!snap) return { ...base, dataSource: childDataSource };
  return {
    ...base,
    dataSource: childDataSource,
    ...(snap.listView !== undefined && { listView: snap.listView }),
    ...(snap.pagination !== undefined && { pagination: snap.pagination }),
    ...(snap.tableConfig !== undefined && { tableConfig: snap.tableConfig }),
    ...(snap.pdfTemplate !== undefined && { pdfTemplate: snap.pdfTemplate }),
  };
}

export function effectiveConfigForListPageBlock(
  config: IDynamicViewConfig,
  block: IListPageBlock
): IDynamicViewConfig {
  const b = block.linkedListBinding;
  if (!b?.listTitle?.trim() || !b?.parentLookupFieldInternalName?.trim()) return config;
  const childDataSource: IDataSourceConfig = { kind: 'list', title: b.listTitle.trim() };
  return mergeLinkedListMemoryIntoConfig(config, childDataSource);
}

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
    const columnsAligned = mergeExtraListPageColumnsIntoLayout(
      layout,
      columnsRaw.map((c) => (Array.isArray(c) ? (c as IListPageBlock[]) : []))
    );
    const columns: IListPageBlock[][] = [];
    for (let ci = 0; ci < need; ci++) {
      const colSrc = columnsAligned[ci];
      const blocks: IListPageBlock[] = [];
      if (Array.isArray(colSrc)) {
        for (let bi = 0; bi < colSrc.length; bi++) {
          const b = colSrc[bi];
          if (!b || typeof b !== 'object') continue;
          const bb = b as unknown as Record<string, unknown>;
          const bid = typeof bb.id === 'string' && bb.id.trim() ? bb.id.trim() : newBlockId();
          const bt = String(bb.type ?? '');
          if (!VALID_BLOCK_TYPES.has(bt)) continue;
          const type = bt as IListPageBlock['type'];
          const linkedListBinding = sanitizeLinkedListBinding(bb.linkedListBinding, type);
          const bindOpt = linkedListBinding ? { linkedListBinding } : {};
          const nestedDash = type === 'dashboard' ? sanitizeBlockDashboard(bb.dashboard) : undefined;
          const pairedListBlockId =
            type === 'dashboard' &&
            typeof bb.pairedListBlockId === 'string' &&
            bb.pairedListBlockId.trim().length > 0
              ? bb.pairedListBlockId.trim()
              : undefined;
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
          if (type === 'sectionTitle') {
            blocks.push({
              id: bid,
              type,
              sectionTitle: sanitizeSectionTitleConfig(bb.sectionTitle),
            });
            continue;
          }
          if (type === 'alert') {
            blocks.push({
              id: bid,
              type,
              ...bindOpt,
              alert: sanitizeAlertConfig(bb.alert),
            });
            continue;
          }
          if (type === 'buttons') {
            blocks.push({
              id: bid,
              type,
              buttons: sanitizeButtonsConfig(bb.buttons),
            });
            continue;
          }
          blocks.push({
            id: bid,
            type,
            ...bindOpt,
            ...(nestedDash ? { dashboard: nestedDash } : {}),
            ...(pairedListBlockId !== undefined ? { pairedListBlockId } : {}),
          });
        }
      }
      columns.push(blocks);
    }
    sections.push({ id, layout, columns });
  }
  if (sections.length === 0) return undefined;
  const out: IListPageLayoutConfig = { sections };
  const pad = sanitizeListPageContentPadding(r.contentPadding);
  if (pad) out.contentPadding = pad;
  return out;
}

export function defaultListPageLayoutFromLegacy(config: IDynamicViewConfig): IListPageLayoutConfig {
  return { sections: buildLegacyListPageSections(config) };
}

export function saveDashboardPairedListBlock(
  config: IDynamicViewConfig,
  dashboardBlockId: string,
  pairedListBlockId: string | undefined
): IDynamicViewConfig {
  const layout = config.listPageLayout;
  const block = layout ? findListPageBlockInSections(layout.sections, dashboardBlockId) : null;
  if (!layout || !block || block.type !== 'dashboard') {
    const nextDash: IDashboardConfig = { ...config.dashboard };
    if (pairedListBlockId?.trim()) {
      nextDash.linkedListBlockId = pairedListBlockId.trim();
    } else {
      delete nextDash.linkedListBlockId;
    }
    return { ...config, dashboard: nextDash };
  }
  const nextBlock: IListPageBlock = { ...block };
  if (pairedListBlockId?.trim()) {
    nextBlock.pairedListBlockId = pairedListBlockId.trim();
  } else {
    delete nextBlock.pairedListBlockId;
  }
  return {
    ...config,
    listPageLayout: replaceBlockInListPageLayout(layout, dashboardBlockId, nextBlock),
  };
}
