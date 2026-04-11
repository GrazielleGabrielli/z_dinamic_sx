import type {
  IListPageAlertBlockConfig,
  IListPageAlertCountRule,
  IListPageBannerBlockConfig,
  IListPageRichEditorBlockConfig,
  IListPageSectionTitleBlockConfig,
  TListPageAlertCountOp,
  TListPageAlertVariant,
  TListPageBannerContentAlign,
  TListPageSectionTitleSize,
} from '../config/types';

export function defaultBannerConfig(): IListPageBannerBlockConfig {
  return {
    imageUrl: '',
    title: '',
    subtitle: '',
    linkUrl: '',
    openInNewTab: false,
    imageAlt: '',
    contentAlign: 'center',
    heightPx: 220,
    overlayOpacity: 0.35,
    showButton: true,
    buttonText: 'Saiba mais',
  };
}

export function defaultRichEditorConfig(): IListPageRichEditorBlockConfig {
  return {
    title: '',
    html: '',
    placeholder: 'Escreva o conteúdo…',
    minHeightPx: 120,
    readOnly: false,
    allowImages: true,
    allowLinks: true,
    allowTables: true,
    allowLists: true,
    allowHeaders: true,
    allowVideoEmbed: false,
  };
}

export function defaultSectionTitleConfig(): IListPageSectionTitleBlockConfig {
  return {
    title: '',
    subtitle: '',
    iconName: '',
    align: 'left',
    showDivider: true,
    size: 'md',
    marginTopPx: 0,
    marginBottomPx: 16,
  };
}

export function defaultAlertConfig(): IListPageAlertBlockConfig {
  return {
    title: '',
    message: '',
    variant: 'info',
    iconName: '',
    dismissible: false,
    emphasized: true,
    linkUrl: '',
    linkText: '',
  };
}

const ALIGN: TListPageBannerContentAlign[] = ['left', 'center', 'right'];
const SECTION_TITLE_SIZES: TListPageSectionTitleSize[] = ['sm', 'md', 'lg'];
const ALERT_VARIANTS: TListPageAlertVariant[] = ['info', 'success', 'warning', 'error'];
const ALERT_COUNT_OPS: TListPageAlertCountOp[] = ['eq', 'ne', 'gt', 'ge', 'lt', 'le'];
export const MAX_ALERT_COUNT_RULES = 20;

export function listAlertCountMatches(actual: number, op: TListPageAlertCountOp, expected: number): boolean {
  switch (op) {
    case 'eq':
      return actual === expected;
    case 'ne':
      return actual !== expected;
    case 'gt':
      return actual > expected;
    case 'ge':
      return actual >= expected;
    case 'lt':
      return actual < expected;
    case 'le':
      return actual <= expected;
    default:
      return false;
  }
}

export function mergeAlertWithCountRule(
  base: IListPageAlertBlockConfig,
  rule: IListPageAlertCountRule
): IListPageAlertBlockConfig {
  const out: IListPageAlertBlockConfig = { ...base };
  if (rule.title !== undefined) out.title = rule.title;
  if (rule.message !== undefined) out.message = rule.message;
  if (rule.variant !== undefined) out.variant = rule.variant;
  if (rule.iconName !== undefined) out.iconName = rule.iconName;
  if (rule.dismissible !== undefined) out.dismissible = rule.dismissible;
  if (rule.emphasized !== undefined) out.emphasized = rule.emphasized;
  if (rule.linkUrl !== undefined) out.linkUrl = rule.linkUrl;
  if (rule.linkText !== undefined) out.linkText = rule.linkText;
  return out;
}

function clamp(n: number, min: number, max: number): number {
  if (isNaN(n)) return min;
  return Math.min(max, Math.max(min, n));
}

export function sanitizeBannerConfig(raw: unknown): IListPageBannerBlockConfig {
  const d = defaultBannerConfig();
  if (!raw || typeof raw !== 'object') return d;
  const o = raw as Record<string, unknown>;
  if (typeof o.imageUrl === 'string') d.imageUrl = o.imageUrl.trim();
  if (typeof o.title === 'string') d.title = o.title;
  if (typeof o.subtitle === 'string') d.subtitle = o.subtitle;
  if (typeof o.linkUrl === 'string') d.linkUrl = o.linkUrl.trim();
  if (o.openInNewTab === true) d.openInNewTab = true;
  if (typeof o.imageAlt === 'string') d.imageAlt = o.imageAlt;
  const al = typeof o.contentAlign === 'string' ? o.contentAlign : '';
  if (ALIGN.indexOf(al as TListPageBannerContentAlign) !== -1) {
    d.contentAlign = al as TListPageBannerContentAlign;
  }
  if (typeof o.heightPx === 'number') d.heightPx = clamp(Math.round(o.heightPx), 80, 800);
  if (typeof o.overlayOpacity === 'number') d.overlayOpacity = clamp(o.overlayOpacity, 0, 1);
  if (o.showButton === false) d.showButton = false;
  if (typeof o.buttonText === 'string') d.buttonText = o.buttonText;
  return d;
}

export function sanitizeRichEditorConfig(raw: unknown): IListPageRichEditorBlockConfig {
  const d = defaultRichEditorConfig();
  if (!raw || typeof raw !== 'object') return d;
  const o = raw as Record<string, unknown>;
  if (typeof o.title === 'string') d.title = o.title;
  if (typeof o.html === 'string') d.html = o.html;
  if (typeof o.placeholder === 'string') d.placeholder = o.placeholder;
  if (typeof o.minHeightPx === 'number') d.minHeightPx = clamp(Math.round(o.minHeightPx), 40, 2000);
  if (o.readOnly === true) d.readOnly = true;
  if (o.allowImages === false) d.allowImages = false;
  if (o.allowLinks === false) d.allowLinks = false;
  if (o.allowTables === false) d.allowTables = false;
  if (o.allowLists === false) d.allowLists = false;
  if (o.allowHeaders === false) d.allowHeaders = false;
  if (o.allowVideoEmbed === true) d.allowVideoEmbed = true;
  return d;
}

export function sanitizeSectionTitleConfig(raw: unknown): IListPageSectionTitleBlockConfig {
  const d = defaultSectionTitleConfig();
  if (!raw || typeof raw !== 'object') return d;
  const o = raw as Record<string, unknown>;
  if (typeof o.title === 'string') d.title = o.title;
  if (typeof o.subtitle === 'string') d.subtitle = o.subtitle;
  if (typeof o.iconName === 'string') d.iconName = o.iconName.trim();
  const al = typeof o.align === 'string' ? o.align : '';
  if (ALIGN.indexOf(al as TListPageBannerContentAlign) !== -1) {
    d.align = al as TListPageBannerContentAlign;
  }
  if (o.showDivider === false) d.showDivider = false;
  const sz = typeof o.size === 'string' ? o.size : '';
  if (SECTION_TITLE_SIZES.indexOf(sz as TListPageSectionTitleSize) !== -1) {
    d.size = sz as TListPageSectionTitleSize;
  }
  if (typeof o.marginTopPx === 'number') d.marginTopPx = clamp(Math.round(o.marginTopPx), 0, 120);
  if (typeof o.marginBottomPx === 'number') d.marginBottomPx = clamp(Math.round(o.marginBottomPx), 0, 120);
  return d;
}

function sanitizeAlertCountRule(raw: unknown): IListPageAlertCountRule | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const r = raw as Record<string, unknown>;
  const id = typeof r.id === 'string' ? r.id.trim() : '';
  if (!id) return undefined;
  const opRaw = typeof r.countOp === 'string' ? r.countOp : '';
  const countOp: TListPageAlertCountOp = ALERT_COUNT_OPS.indexOf(opRaw as TListPageAlertCountOp) !== -1
    ? (opRaw as TListPageAlertCountOp)
    : 'eq';
  const count = typeof r.count === 'number' && isFinite(r.count) ? Math.round(r.count) : NaN;
  if (!isFinite(count)) return undefined;
  const odataFilter = typeof r.odataFilter === 'string' ? r.odataFilter.trim() : '';
  const out: IListPageAlertCountRule = { id, countOp, count, ...(odataFilter ? { odataFilter } : {}) };
  if (typeof r.title === 'string') out.title = r.title;
  if (typeof r.message === 'string') out.message = r.message;
  const v = typeof r.variant === 'string' ? r.variant : '';
  if (ALERT_VARIANTS.indexOf(v as TListPageAlertVariant) !== -1) {
    out.variant = v as TListPageAlertVariant;
  }
  if (typeof r.iconName === 'string') out.iconName = r.iconName.trim();
  if (r.dismissible === true) out.dismissible = true;
  if (r.dismissible === false) out.dismissible = false;
  if (r.emphasized === true) out.emphasized = true;
  if (r.emphasized === false) out.emphasized = false;
  if (typeof r.linkUrl === 'string') out.linkUrl = r.linkUrl.trim();
  if (typeof r.linkText === 'string') out.linkText = r.linkText;
  return out;
}

export function sanitizeAlertConfig(raw: unknown): IListPageAlertBlockConfig {
  const d = defaultAlertConfig();
  if (!raw || typeof raw !== 'object') return d;
  const o = raw as Record<string, unknown>;
  if (typeof o.title === 'string') d.title = o.title;
  if (typeof o.message === 'string') d.message = o.message;
  const v = typeof o.variant === 'string' ? o.variant : '';
  if (ALERT_VARIANTS.indexOf(v as TListPageAlertVariant) !== -1) {
    d.variant = v as TListPageAlertVariant;
  }
  if (typeof o.iconName === 'string') d.iconName = o.iconName.trim();
  if (o.dismissible === true) d.dismissible = true;
  if (o.emphasized === false) d.emphasized = false;
  if (typeof o.linkUrl === 'string') d.linkUrl = o.linkUrl.trim();
  if (typeof o.linkText === 'string') d.linkText = o.linkText;
  const rulesRaw = Array.isArray(o.countRules) ? o.countRules : [];
  const rules: IListPageAlertCountRule[] = [];
  for (let i = 0; i < rulesRaw.length && rules.length < MAX_ALERT_COUNT_RULES; i++) {
    const rr = sanitizeAlertCountRule(rulesRaw[i]);
    if (rr) rules.push(rr);
  }
  if (rules.length) {
    (d as IListPageAlertBlockConfig).countRules = rules;
  }
  return d;
}
