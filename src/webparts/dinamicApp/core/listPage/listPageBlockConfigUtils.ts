import type {
  IListPageAlertBlockConfig,
  IListPageBannerBlockConfig,
  IListPageRichEditorBlockConfig,
  IListPageSectionTitleBlockConfig,
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
  return d;
}
