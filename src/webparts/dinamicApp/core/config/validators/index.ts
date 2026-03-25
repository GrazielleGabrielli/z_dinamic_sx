import {
  IDynamicViewConfig,
  IDataSourceConfig,
  IDashboardConfig,
  IDashboardCardConfig,
  IDashboardCardStyleConfig,
  IPaginationConfig,
  IListViewConfig,
  IPdfTemplateConfig,
  IPdfTemplateElement,
} from '../types';
import { getDefaultConfig } from '../utils';
import { isValidPdfPageFormat } from '../../pdf/pdfPageFormats';
import { sanitizeTableCssSlots } from '../../../components/DataTable/tableLayoutClasses';

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
  return (
    typeof d.enabled === 'boolean' &&
    typeof d.cardsCount === 'number' &&
    Array.isArray(d.cards) &&
    (d.cards as unknown[]).every(isValidCard)
  );
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
  if (l.sort !== null && (typeof l.sort !== 'object' || !('field' in (l.sort as object)) || !('ascending' in (l.sort as object)))) return false;
  return true;
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

function isValidPdfTemplate(t: unknown): t is IPdfTemplateConfig {
  if (!t || typeof t !== 'object') return false;
  const c = t as Record<string, unknown>;
  if (!isValidPdfPageFormat(c.pageFormat)) return false;
  if (c.orientation !== 'portrait' && c.orientation !== 'landscape') return false;
  if (c.header !== undefined && !isValidPdfSection(c.header)) return false;
  if (c.footer !== undefined && !isValidPdfSection(c.footer)) return false;
  if (!c.body || !isValidPdfSection(c.body)) return false;
  return true;
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

export function parseConfig(raw: string | undefined): IDynamicViewConfig | undefined {
  if (!raw) return undefined;
  try {
    const parsed: unknown = JSON.parse(raw);
    if (!isValidConfig(parsed)) return undefined;
    const defaults = getDefaultConfig();
    const c = parsed as IDynamicViewConfig;
    if (c.listView === undefined) {
      return { ...c, listView: defaults.listView };
    }
    const lv = c.listView;
    const cssSlots = sanitizeTableCssSlots(lv.customTableCssSlots);
    return {
      ...c,
      listView: {
        columns: lv.columns ?? defaults.listView.columns,
        filters: lv.filters ?? defaults.listView.filters,
        sort: lv.sort ?? defaults.listView.sort,
        viewModes: lv.viewModes ?? defaults.listView.viewModes,
        activeViewModeId: lv.activeViewModeId ?? defaults.listView.activeViewModeId,
        pdfExportEnabled: lv.pdfExportEnabled ?? false,
        ...(cssSlots ? { customTableCssSlots: cssSlots } : {}),
        ...(typeof lv.customTableCss === 'string' ? { customTableCss: lv.customTableCss } : {}),
      },
      ...(isValidPdfTemplate(c.pdfTemplate) && { pdfTemplate: c.pdfTemplate }),
    };
  } catch {
    return undefined;
  }
}
