import type {
  IListPageBannerBlockConfig,
  IListPageRichEditorBlockConfig,
  TListPageBannerContentAlign,
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

const ALIGN: TListPageBannerContentAlign[] = ['left', 'center', 'right'];

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
