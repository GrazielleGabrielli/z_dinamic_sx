import type {
  IListPageAlertBlockConfig,
  IListPageAlertCountRule,
  IListPageBannerBlockConfig,
  IListPageButtonItemConfig,
  IListPageButtonsBlockConfig,
  IListPageRichEditorBlockConfig,
  IListPageSectionTitleBlockConfig,
  TListPageAlertCountOp,
  TListPageAlertCountFilterFieldOp,
  TListPageAlertVariant,
  TListPageBannerContentAlign,
  TListPageButtonActionKind,
  TListPageButtonVariant,
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

export const MAX_LIST_PAGE_BUTTONS = 20;

function newListPageButtonItemId(): string {
  return `lpbtn_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
}

export function defaultButtonsConfig(): IListPageButtonsBlockConfig {
  return {
    items: [
      {
        id: newListPageButtonItemId(),
        label: 'Recarregar',
        actionKind: 'reload',
      },
    ],
  };
}

function safeListPageButtonRedirectUrl(raw: string): string {
  const t = raw.trim();
  if (!t) return '';
  if (/^javascript:/i.test(t)) return '';
  return t;
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

const ALERT_COUNT_FILTER_FIELD_OPS: TListPageAlertCountFilterFieldOp[] = [
  'eq',
  'ne',
  'gt',
  'ge',
  'lt',
  'le',
  'contains',
];

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
  const cf = typeof r.countFilterField === 'string' ? r.countFilterField.trim() : '';
  if (cf) {
    out.countFilterField = cf;
    const cfOpRaw = typeof r.countFilterFieldOp === 'string' ? r.countFilterFieldOp : '';
    out.countFilterFieldOp =
      ALERT_COUNT_FILTER_FIELD_OPS.indexOf(cfOpRaw as TListPageAlertCountFilterFieldOp) !== -1
        ? (cfOpRaw as TListPageAlertCountFilterFieldOp)
        : 'eq';
    if (typeof r.countFilterValue === 'string') out.countFilterValue = r.countFilterValue;
  }
  if (r.countFilterUseManualOdata === true) out.countFilterUseManualOdata = true;
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

export function sanitizeButtonsConfig(raw: unknown): IListPageButtonsBlockConfig {
  const fallback = defaultButtonsConfig();
  if (!raw || typeof raw !== 'object') return fallback;
  const o = raw as Record<string, unknown>;
  const itemsRaw = Array.isArray(o.items) ? o.items : [];
  const items: IListPageButtonItemConfig[] = [];
  for (let i = 0; i < itemsRaw.length && items.length < MAX_LIST_PAGE_BUTTONS; i++) {
    const e = itemsRaw[i];
    if (!e || typeof e !== 'object') continue;
    const r = e as Record<string, unknown>;
    const id = typeof r.id === 'string' && r.id.trim() ? r.id.trim() : newListPageButtonItemId();
    const label = typeof r.label === 'string' ? r.label.trim() : '';
    if (!label) continue;
    const akRaw = r.actionKind === 'reload' ? 'reload' : r.actionKind === 'redirect' ? 'redirect' : '';
    const actionKind: TListPageButtonActionKind =
      akRaw === 'reload' || akRaw === 'redirect' ? akRaw : 'redirect';
    const variant: TListPageButtonVariant | undefined =
      r.variant === 'primary' ? 'primary' : r.variant === 'default' ? 'default' : undefined;
    const iconName = typeof r.iconName === 'string' && r.iconName.trim() ? r.iconName.trim() : undefined;
    const css = typeof r.css === 'string' && r.css.trim() ? r.css.trim() : undefined;
    if (actionKind === 'reload') {
      items.push({ id, label, actionKind: 'reload', ...(variant ? { variant } : {}), ...(iconName ? { iconName } : {}), ...(css ? { css } : {}) });
      continue;
    }
    const url = typeof r.url === 'string' ? safeListPageButtonRedirectUrl(r.url) : '';
    if (!url) continue;
    items.push({
      id,
      label,
      actionKind: 'redirect',
      url,
      openInNewTab: r.openInNewTab === true,
      ...(variant ? { variant } : {}),
      ...(iconName ? { iconName } : {}),
      ...(css ? { css } : {}),
    });
  }
  if (items.length === 0) return fallback;

  const alignRaw = o.align;
  const align: IListPageButtonsBlockConfig['align'] =
    alignRaw === 'left' || alignRaw === 'center' || alignRaw === 'right' ? alignRaw : undefined;
  const gapRaw = typeof o.gap === 'number' ? o.gap : undefined;
  const gap = gapRaw !== undefined && gapRaw >= 0 && gapRaw <= 120 ? Math.round(gapRaw) : undefined;
  const containerCss = typeof o.containerCss === 'string' && o.containerCss.trim() ? o.containerCss.trim() : undefined;

  return {
    items,
    ...(align ? { align } : {}),
    ...(gap !== undefined ? { gap } : {}),
    ...(containerCss ? { containerCss } : {}),
  };
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
