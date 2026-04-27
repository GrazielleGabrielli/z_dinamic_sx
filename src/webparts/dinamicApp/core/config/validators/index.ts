import {
  IDynamicViewConfig,
  IDataSourceConfig,
  IDashboardConfig,
  IDashboardCardConfig,
  IDashboardCardStyleConfig,
  IProjectManagementConfig,
  IPaginationConfig,
  IListViewConfig,
  IListRowActionConfig,
  IPdfTemplateConfig,
  IPdfTemplateElement,
  TViewMode,
} from '../types';
import { getDefaultConfig } from '../utils';
import { isValidPdfPageFormat } from '../../pdf/pdfPageFormats';
import { sanitizeTableCssSlots } from '../../../components/DataTable/tableLayoutClasses';
import { sanitizeTableRowStyleRules } from '../../table/utils/tableRowStyleRuleEval';
import {
  normalizeListPageLayoutDashboards,
  sanitizeListPageLayout,
} from '../../listPage/listPageLayoutUtils';
import { sanitizeFormManagerConfig } from '../../formManager/sanitizeFormManagerConfig';

const VALID_MODES = ['list', 'projectManagement', 'formManager'];
const VALID_AGGREGATES = ['count', 'sum'];
const VALID_VARIANTS = ['default', 'outlined', 'soft', 'solid'];
const VALID_BORDER_RADIUS = ['none', 'sm', 'md', 'lg', 'xl', 'full'];
const VALID_PADDING = ['sm', 'md', 'lg'];
const VALID_SHADOW = ['none', 'sm', 'md', 'lg'];
const VALID_TITLE_SIZE = ['xs', 'sm', 'md', 'lg'];
const VALID_SUBTITLE_SIZE = ['xs', 'sm', 'md'];
const VALID_VALUE_SIZE = ['lg', 'xl', '2xl', '3xl'];
const VALID_FONT_WEIGHT = ['normal', 'medium', 'semibold', 'bold'];
const VALID_ALIGN = ['left', 'center', 'right'];
const VALID_ICON_POSITION = ['left', 'top', 'right'];
const VALID_LOADING_STYLE = ['skeleton', 'spinner', 'text'];

function isValidDataSource(ds: unknown): ds is IDataSourceConfig {
  if (!ds || typeof ds !== 'object') return false;
  const d = ds as Record<string, unknown>;
  if (d.webServerRelativeUrl !== undefined && typeof d.webServerRelativeUrl !== 'string') return false;
  return (
    (d.kind === 'list' || d.kind === 'library') &&
    typeof d.title === 'string' &&
    (d.title as string).trim().length > 0
  );
}

function isValidCardStyle(s: unknown): s is IDashboardCardStyleConfig {
  if (!s || typeof s !== 'object') return false;
  const st = s as Record<string, unknown>;
  if (st.variant !== undefined && VALID_VARIANTS.indexOf(st.variant as string) === -1) return false;
  if (st.borderRadius !== undefined && VALID_BORDER_RADIUS.indexOf(st.borderRadius as string) === -1) return false;
  if (st.padding !== undefined && VALID_PADDING.indexOf(st.padding as string) === -1) return false;
  if (st.shadow !== undefined && VALID_SHADOW.indexOf(st.shadow as string) === -1) return false;
  if (st.titleSize !== undefined && VALID_TITLE_SIZE.indexOf(st.titleSize as string) === -1) return false;
  if (st.subtitleSize !== undefined && VALID_SUBTITLE_SIZE.indexOf(st.subtitleSize as string) === -1) return false;
  if (st.valueSize !== undefined && VALID_VALUE_SIZE.indexOf(st.valueSize as string) === -1) return false;
  if (st.titleWeight !== undefined && VALID_FONT_WEIGHT.indexOf(st.titleWeight as string) === -1) return false;
  if (st.valueWeight !== undefined && VALID_FONT_WEIGHT.indexOf(st.valueWeight as string) === -1) return false;
  if (st.align !== undefined && VALID_ALIGN.indexOf(st.align as string) === -1) return false;
  if (st.iconPosition !== undefined && VALID_ICON_POSITION.indexOf(st.iconPosition as string) === -1) return false;
  if (st.loadingStyle !== undefined && VALID_LOADING_STYLE.indexOf(st.loadingStyle as string) === -1) return false;
  if (st.showIcon === true && (st.iconName === undefined || typeof st.iconName !== 'string' || (st.iconName as string).trim() === '')) return false;
  return true;
}

function isValidCard(card: unknown): card is IDashboardCardConfig {
  if (!card || typeof card !== 'object') return false;
  const c = card as Record<string, unknown>;
  if (typeof c.id !== 'string' || typeof c.title !== 'string') return false;
  if (typeof c.aggregate !== 'string' || VALID_AGGREGATES.indexOf(c.aggregate as string) === -1) return false;
  if ((c.aggregate as string) === 'sum' && (!c.field || typeof c.field !== 'string' || (c.field as string).trim() === '')) return false;
  if (c.subtitle !== undefined && typeof c.subtitle !== 'string') return false;
  if (c.emptyValueText !== undefined && typeof c.emptyValueText !== 'string') return false;
  if (c.errorText !== undefined && typeof c.errorText !== 'string') return false;
  if (c.loadingText !== undefined && typeof c.loadingText !== 'string') return false;
  if (c.style !== undefined && !isValidCardStyle(c.style)) return false;
  return true;
}

function isValidDashboard(db: unknown): db is IDashboardConfig {
  if (!db || typeof db !== 'object') return false;
  const d = db as Record<string, unknown>;
  if (typeof d.enabled !== 'boolean') return false;
  const cards = Array.isArray(d.cards) ? (d.cards as unknown[]) : [];
  if (!cards.every(isValidCard)) return false;
  const cnt =
    typeof d.cardsCount === 'number' && !Number.isNaN(d.cardsCount as number)
      ? (d.cardsCount as number)
      : cards.length;
  if (cnt < 0) return false;
  return true;
}

/** Garante `cards`/`cardsCount` e cópia de `chartSeries` para JSON válido e runtime seguro. */
export function coerceDashboardShape(d: IDashboardConfig): IDashboardConfig {
  const defaults = getDefaultConfig().dashboard;
  const cards = Array.isArray(d.cards) ? d.cards : [];
  const out: IDashboardConfig = {
    ...defaults,
    ...d,
    cards,
    cardsCount:
      typeof d.cardsCount === 'number' && !Number.isNaN(d.cardsCount) ? d.cardsCount : cards.length,
  };
  if (Array.isArray(d.chartSeries)) {
    out.chartSeries = d.chartSeries.map((s) => ({ ...s }));
  }
  return out;
}

function isValidPagination(pg: unknown): pg is IPaginationConfig {
  if (!pg || typeof pg !== 'object') return false;
  const p = pg as Record<string, unknown>;
  return (
    typeof p.enabled === 'boolean' &&
    typeof p.pageSize === 'number' &&
    Array.isArray(p.pageSizeOptions)
  );
}

function isValidListView(lv: unknown): lv is IListViewConfig {
  if (!lv || typeof lv !== 'object') return false;
  const l = lv as Record<string, unknown>;
  if (!Array.isArray(l.columns) || !Array.isArray(l.filters)) return false;
  if (l.sort != null && (typeof l.sort !== 'object' || !('field' in (l.sort as object)) || !('ascending' in (l.sort as object)))) return false;
  return true;
}

function sanitizeListRowActions(raw: unknown): IListRowActionConfig[] | undefined {
  if (!Array.isArray(raw)) return undefined;
  const presets = new Set(['view', 'edit', 'link', 'custom']);
  const out: IListRowActionConfig[] = [];
  for (let i = 0; i < raw.length; i++) {
    const entry = raw[i];
    if (!entry || typeof entry !== 'object') continue;
    const e = entry as Record<string, unknown>;
    const id = typeof e.id === 'string' ? e.id.trim() : '';
    const title = typeof e.title === 'string' ? e.title.trim() : '';
    const urlTemplate = typeof e.urlTemplate === 'string' ? e.urlTemplate.trim() : '';
    if (!id || !title || !urlTemplate) continue;
    const iconKey = typeof e.iconPreset === 'string' ? e.iconPreset : 'link';
    const iconPreset = presets.has(iconKey) ? (iconKey as IListRowActionConfig['iconPreset']) : 'link';
    const scope = e.scope === 'wholeRow' ? 'wholeRow' : 'icon';
    const customIconName =
      typeof e.customIconName === 'string' && e.customIconName.trim() ? e.customIconName.trim() : undefined;
    out.push({
      id,
      title,
      iconPreset,
      ...(customIconName ? { customIconName } : {}),
      urlTemplate,
      openInNewTab: e.openInNewTab === true,
      scope,
    });
  }
  return out.length > 0 ? out : undefined;
}

function sanitizeProjectManagementConfig(raw: unknown): IProjectManagementConfig | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const src = raw as Record<string, unknown>;
  const rawColumns = Array.isArray(src.columns) ? src.columns : [];
  const columns: IProjectManagementConfig['columns'] = [];
  for (let i = 0; i < rawColumns.length; i++) {
    const entry = rawColumns[i];
    if (!entry || typeof entry !== 'object') continue;
    const e = entry as Record<string, unknown>;
    const id = typeof e.id === 'string' ? e.id.trim() : '';
    const title = typeof e.title === 'string' ? e.title.trim() : '';
    const rawRules = Array.isArray(e.rules) ? e.rules : [];
    const rules: NonNullable<IProjectManagementConfig['columns']>[number]['rules'] = [];
    for (let j = 0; j < rawRules.length; j++) {
      const rr = rawRules[j];
      if (!rr || typeof rr !== 'object') continue;
      const r = rr as Record<string, unknown>;
      const ruleId = typeof r.id === 'string' ? r.id.trim() : `rule_${i + 1}_${j + 1}`;
      const field = typeof r.field === 'string' ? r.field.trim() : '';
      const value = typeof r.value === 'string' ? r.value : '';
      if (!field) continue;
      rules.push({ id: ruleId, field, value });
    }
    if (!id || !title) continue;
    columns.push({
      id,
      title,
      rules,
    });
  }
  return { columns };
}

function isValidPdfElement(el: unknown): el is IPdfTemplateElement {
  if (!el || typeof el !== 'object') return false;
  const e = el as Record<string, unknown>;
  return (
    typeof e.id === 'string' &&
    typeof e.type === 'string' &&
    ['text', 'image', 'rect', 'line'].indexOf(e.type as string) !== -1 &&
    typeof e.x === 'number' &&
    typeof e.y === 'number'
  );
}

function isValidPdfSection(s: unknown): boolean {
  if (!s || typeof s !== 'object') return false;
  const sec = s as Record<string, unknown>;
  if (!Array.isArray(sec.elements)) return false;
  return (sec.elements as unknown[]).every(isValidPdfElement);
}

export function isValidPdfTemplate(t: unknown): t is IPdfTemplateConfig {
  if (!t || typeof t !== 'object') return false;
  const c = t as Record<string, unknown>;
  if (!isValidPdfPageFormat(c.pageFormat)) return false;
  if (c.orientation !== 'portrait' && c.orientation !== 'landscape') return false;
  if (c.header !== undefined && !isValidPdfSection(c.header)) return false;
  if (c.footer !== undefined && !isValidPdfSection(c.footer)) return false;
  if (!c.body || !isValidPdfSection(c.body)) return false;
  return true;
}

const VALID_PAGINATION_LAYOUTS = new Set(['buttons', 'numbered', 'compact', 'paged']);

export function sanitizeListViewConfig(lv: unknown): IListViewConfig | undefined {
  if (!lv || typeof lv !== 'object' || !isValidListView(lv)) return undefined;
  const defaults = getDefaultConfig().listView;
  const lvo = lv as IListViewConfig;
  const cssSlots = sanitizeTableCssSlots(lvo.customTableCssSlots);
  const rowRules = sanitizeTableRowStyleRules(lvo.tableRowStyleRules);
  const listRowActions = sanitizeListRowActions(lvo.listRowActions);
  return {
    columns: lvo.columns ?? defaults.columns,
    filters: lvo.filters ?? defaults.filters,
    sort: lvo.sort ?? defaults.sort,
    viewModes: lvo.viewModes ?? defaults.viewModes,
    activeViewModeId: lvo.activeViewModeId ?? defaults.activeViewModeId,
    pdfExportEnabled: lvo.pdfExportEnabled ?? false,
    ...(lvo.listCardViewEnabled === true ? { listCardViewEnabled: true } : {}),
    ...(lvo.listCardViewEnabled === true && lvo.listDefaultDisplayMode === 'cards'
      ? { listDefaultDisplayMode: 'cards' as const }
      : {}),
    ...(cssSlots ? { customTableCssSlots: cssSlots } : {}),
    ...(typeof lvo.customTableCss === 'string' ? { customTableCss: lvo.customTableCss } : {}),
    ...(rowRules ? { tableRowStyleRules: rowRules } : {}),
    ...(listRowActions ? { listRowActions } : {}),
  };
}

export interface IListTableEditorBundle {
  listView: IListViewConfig;
  pagination: IPaginationConfig;
  pdfTemplate?: IPdfTemplateConfig;
  projectManagement?: IProjectManagementConfig;
}

export function sanitizeListTableEditorBundle(
  raw: unknown,
  fallback: IListTableEditorBundle,
  mode: TViewMode
): IListTableEditorBundle | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const o = raw as Record<string, unknown>;
  let listView = fallback.listView;
  if (o.listView !== undefined) {
    const s = sanitizeListViewConfig(o.listView);
    if (!s) return undefined;
    listView = s;
  }
  let pagination = fallback.pagination;
  if (o.pagination !== undefined) {
    if (!o.pagination || typeof o.pagination !== 'object') return undefined;
    const p = o.pagination as Record<string, unknown>;
    const layoutRaw = typeof p.layout === 'string' ? p.layout : '';
    const layout =
      VALID_PAGINATION_LAYOUTS.has(layoutRaw) ? (layoutRaw as IPaginationConfig['layout']) : fallback.pagination.layout ?? 'buttons';
    const pageSizeOptions = Array.isArray(p.pageSizeOptions)
      ? (p.pageSizeOptions as unknown[])
          .filter((n): n is number => typeof n === 'number' && n > 0)
          .slice(0, 12)
      : fallback.pagination.pageSizeOptions;
    const pageSize =
      typeof p.pageSize === 'number'
        ? Math.min(500, Math.max(1, Math.round(p.pageSize)))
        : fallback.pagination.pageSize;
    pagination = {
      ...fallback.pagination,
      enabled: typeof p.enabled === 'boolean' ? p.enabled : fallback.pagination.enabled,
      pageSize,
      layout,
      pageSizeOptions: pageSizeOptions.length ? pageSizeOptions : fallback.pagination.pageSizeOptions,
    };
  }
  let pdfTemplate = fallback.pdfTemplate;
  if (o.pdfTemplate !== undefined) {
    if (o.pdfTemplate === null) {
      pdfTemplate = undefined;
    } else if (isValidPdfTemplate(o.pdfTemplate)) {
      pdfTemplate = o.pdfTemplate;
    } else {
      return undefined;
    }
  }
  let projectManagement = fallback.projectManagement;
  if (mode === 'projectManagement' && o.projectManagement !== undefined) {
    const pm = sanitizeProjectManagementConfig(o.projectManagement);
    projectManagement = pm ?? { columns: [] };
  }
  return { listView, pagination, pdfTemplate, projectManagement };
}

export function isValidConfig(config: unknown): config is IDynamicViewConfig {
  if (!config || typeof config !== 'object') return false;
  const c = config as Record<string, unknown>;
  const base =
    isValidDataSource(c.dataSource) &&
    typeof c.mode === 'string' &&
    VALID_MODES.indexOf(c.mode as string) !== -1 &&
    isValidDashboard(c.dashboard) &&
    isValidPagination(c.pagination);
  if (!base) return false;
  if (c.listView !== undefined && !isValidListView(c.listView)) return false;
  if (c.pdfTemplate !== undefined && !isValidPdfTemplate(c.pdfTemplate)) return false;
  return true;
}

/**
 * Se o JSON falhar validação mas for modo formulário com fonte de dados válida,
 * repõe partes legadas (dashboard, paginação, listView) com defaults para não perder `formManager`.
 */
function repairParsedConfigForFormManagerIfNeeded(parsed: unknown): unknown {
  if (isValidConfig(parsed)) return parsed;
  if (!parsed || typeof parsed !== 'object') return parsed;
  const o = parsed as Record<string, unknown>;
  if (o.mode !== 'formManager') return parsed;
  if (!isValidDataSource(o.dataSource)) return parsed;
  const defaults = getDefaultConfig();
  let next: Record<string, unknown> = { ...o };
  if (!isValidDashboard(next.dashboard)) {
    next = { ...next, dashboard: defaults.dashboard };
  }
  if (!isValidPagination(next.pagination)) {
    next = { ...next, pagination: defaults.pagination };
  }
  if (next.listView !== undefined && !isValidListView(next.listView)) {
    next = { ...next, listView: defaults.listView };
  }
  return next;
}

export function parseConfig(raw: string | undefined): IDynamicViewConfig | undefined {
  if (!raw) return undefined;
  try {
    const parsed: unknown = JSON.parse(raw);
    const candidate = repairParsedConfigForFormManagerIfNeeded(parsed);
    if (!isValidConfig(candidate)) return undefined;
    const defaults = getDefaultConfig();
    const c = candidate as IDynamicViewConfig;
    const projectManagement = sanitizeProjectManagementConfig(c.projectManagement);
    const formManager = sanitizeFormManagerConfig((c as unknown as Record<string, unknown>).formManager);
    if (c.listView === undefined) {
      const listPageLayoutEarly = sanitizeListPageLayout(
        (c as unknown as Record<string, unknown>).listPageLayout
      );
      const listPageLayoutNorm =
        listPageLayoutEarly !== undefined
          ? normalizeListPageLayoutDashboards(listPageLayoutEarly, c.dashboard)
          : undefined;
      return {
        ...c,
        dashboard: coerceDashboardShape(c.dashboard),
        listView: defaults.listView,
        projectManagement: projectManagement ?? defaults.projectManagement,
        ...(formManager ? { formManager } : {}),
        ...(listPageLayoutNorm ? { listPageLayout: listPageLayoutNorm } : {}),
      };
    }
    const sanitizedListView = sanitizeListViewConfig(c.listView);
    if (!sanitizedListView) return undefined;
    const listPageLayoutRaw = sanitizeListPageLayout((c as unknown as Record<string, unknown>).listPageLayout);
    const listPageLayout =
      listPageLayoutRaw !== undefined
        ? normalizeListPageLayoutDashboards(listPageLayoutRaw, c.dashboard)
        : undefined;
    return {
      ...c,
      dashboard: coerceDashboardShape(c.dashboard),
      ...(formManager ? { formManager } : {}),
      listView: sanitizedListView,
      projectManagement: projectManagement ?? defaults.projectManagement,
      ...(isValidPdfTemplate(c.pdfTemplate) && { pdfTemplate: c.pdfTemplate }),
      ...(listPageLayout ? { listPageLayout } : {}),
    };
  } catch {
    return undefined;
  }
}
