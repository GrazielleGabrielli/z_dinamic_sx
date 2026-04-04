import type { IListPageRichEditorBlockConfig } from '../config/types';

const CORE = new Set([
  'P',
  'DIV',
  'SPAN',
  'BR',
  'STRONG',
  'B',
  'EM',
  'I',
  'U',
  'SUB',
  'SUP',
  'HR',
  'BLOCKQUOTE',
]);

function isSafeHttpUrl(s: string): boolean {
  const t = s.trim().toLowerCase();
  return (
    t.indexOf('https://') === 0 ||
    t.indexOf('http://') === 0 ||
    t.indexOf('mailto:') === 0 ||
    t.indexOf('tel:') === 0
  );
}

function isAllowedIframeSrc(src: string): boolean {
  const t = src.trim();
  try {
    const u = new URL(t);
    if (u.protocol !== 'https:' && u.protocol !== 'http:') return false;
    const h = u.hostname.toLowerCase();
    if (h === 'www.youtube.com' || h === 'youtube.com') return u.pathname.indexOf('/embed/') === 0;
    if (h === 'www.youtube-nocookie.com' || h === 'youtube-nocookie.com')
      return u.pathname.indexOf('/embed/') === 0;
    if (h === 'player.vimeo.com') return u.pathname.length > 1;
    return false;
  } catch {
    return false;
  }
}

function allowedTag(tag: string, cfg: IListPageRichEditorBlockConfig): boolean {
  if (CORE.has(tag)) return true;
  if (cfg.allowLinks && tag === 'A') return true;
  if (cfg.allowImages && tag === 'IMG') return true;
  if (cfg.allowLists && (tag === 'UL' || tag === 'OL' || tag === 'LI')) return true;
  if (cfg.allowTables && ['TABLE', 'THEAD', 'TBODY', 'TR', 'TH', 'TD', 'CAPTION'].indexOf(tag) !== -1)
    return true;
  if (cfg.allowHeaders && /^H[1-6]$/.test(tag)) return true;
  if (cfg.allowVideoEmbed && tag === 'IFRAME') return true;
  return false;
}

function copyAllowedAttrs(el: Element, tag: string, cfg: IListPageRichEditorBlockConfig): string {
  if (tag === 'A' && cfg.allowLinks) {
    const href = el.getAttribute('href') ?? '';
    const safe = isSafeHttpUrl(href) ? href : '#';
    const target = el.getAttribute('target') === '_blank' ? ' target="_blank" rel="noopener noreferrer"' : '';
    return ` href="${safe.replace(/"/g, '&quot;')}"${target}`;
  }
  if (tag === 'IMG' && cfg.allowImages) {
    const src = el.getAttribute('src') ?? '';
    if (!isSafeHttpUrl(src)) return '';
    const alt = (el.getAttribute('alt') ?? '').replace(/"/g, '&quot;');
    return ` src="${src.replace(/"/g, '&quot;')}" alt="${alt}"`;
  }
  if (tag === 'IFRAME' && cfg.allowVideoEmbed) {
    const src = el.getAttribute('src') ?? '';
    if (!isAllowedIframeSrc(src)) return '';
    return ` src="${src.replace(/"/g, '&quot;')}" allowfullscreen="true"`;
  }
  if (['TD', 'TH'].indexOf(tag) !== -1) {
    const cs = el.getAttribute('colspan');
    const rs = el.getAttribute('rowspan');
    let a = '';
    if (cs && /^\d+$/.test(cs)) a += ` colspan="${cs}"`;
    if (rs && /^\d+$/.test(rs)) a += ` rowspan="${rs}"`;
    return a;
  }
  return '';
}

function nodeToHtml(node: Node, cfg: IListPageRichEditorBlockConfig): string {
  if (node.nodeType === Node.TEXT_NODE) {
    const t = node.textContent ?? '';
    return t
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return '';
  const el = node as Element;
  const tag = el.tagName.toUpperCase();
  if (tag === 'SCRIPT' || tag === 'STYLE') return '';
  if (!allowedTag(tag, cfg)) {
    let inner = '';
    for (let c = el.firstChild; c; c = c.nextSibling) inner += nodeToHtml(c, cfg);
    return inner;
  }
  const attrs = copyAllowedAttrs(el, tag, cfg);
  let inner = '';
  for (let c = el.firstChild; c; c = c.nextSibling) inner += nodeToHtml(c, cfg);
  if (tag === 'BR' || tag === 'HR') return `<${tag.toLowerCase()}${attrs} />`;
  if (tag === 'IMG') return `<img${attrs} />`;
  if (tag === 'IFRAME') return `<iframe${attrs} title=""></iframe>`;
  return `<${tag.toLowerCase()}${attrs}>${inner}</${tag.toLowerCase()}>`;
}

export function sanitizeRichEditorHtml(html: string, cfg: IListPageRichEditorBlockConfig): string {
  const trimmed = (html ?? '').trim();
  if (!trimmed) return '';
  if (typeof DOMParser === 'undefined') return '';
  try {
    const doc = new DOMParser().parseFromString(`<div id="root">${trimmed}</div>`, 'text/html');
    const root = doc.getElementById('root');
    if (!root) return '';
    let out = '';
    for (let c = root.firstChild; c; c = c.nextSibling) out += nodeToHtml(c, cfg);
    return out;
  } catch {
    return '';
  }
}
