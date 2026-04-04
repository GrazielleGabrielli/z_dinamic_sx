import type { IDynamicContext } from '../../dynamicTokens/types';
import { resolveValue } from '../../dynamicTokens';

function formatScalar(v: unknown): string {
  if (v == null) return '';
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (Array.isArray(v)) return v.map(formatScalar).filter(Boolean).join(', ');
  if (typeof v === 'object') {
    const o = v as Record<string, unknown>;
    if (o.Title != null) return String(o.Title);
    if (o.title != null) return String(o.title);
    if (o.Id != null) return String(o.Id);
  }
  return '';
}

function readPath(item: Record<string, unknown>, path: string): string {
  const parts = path.split('/').map((p) => p.trim()).filter(Boolean);
  let cur: unknown = item;
  for (let i = 0; i < parts.length; i++) {
    if (cur == null || typeof cur !== 'object') return '';
    cur = (cur as Record<string, unknown>)[parts[i]];
  }
  return formatScalar(cur);
}

/** Resolve campo com caminho `Pai/Filho`; tenta chave case-insensitive no primeiro nível e `Id` para `ID`/`id`. */
function resolveFieldKey(item: Record<string, unknown>, rawKey: string): string {
  const k = rawKey.trim();
  if (!k) return '';
  const fromPath = readPath(item, k);
  if (fromPath !== '') return fromPath;
  const lower = k.toLowerCase();
  if (lower === 'id') {
    const idVal = item.Id ?? item.id;
    if (idVal != null && idVal !== undefined) return String(idVal);
  }
  const keys = Object.keys(item);
  for (let i = 0; i < keys.length; i++) {
    if (keys[i].toLowerCase() === lower) {
      return formatScalar(item[keys[i]]);
    }
  }
  return '';
}

/**
 * Substitui `{{Campo}}`, `{Campo}` / `{Pai/Filho}` e tokens [me], [siteurl], [query:x] na URL.
 * `{{ID}}` e `{{ Id }}` resolvem para o Id do item (SharePoint).
 */
export function resolveListRowActionUrl(
  template: string,
  item: Record<string, unknown>,
  context: IDynamicContext
): string {
  let s = template.replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (_match, rawKey: string) => {
    const val = resolveFieldKey(item, rawKey);
    return encodeURIComponent(val);
  });
  s = s.replace(/\{([^}]+)\}/g, (_match, rawKey: string) => {
    const val = resolveFieldKey(item, rawKey.trim());
    return encodeURIComponent(val);
  });
  s = s.replace(/\[[^\]]+\]/gi, (match) => {
    const resolved = resolveValue(match, context);
    if (resolved === undefined || resolved === null) return '';
    return String(resolved);
  });
  return s.trim();
}

export function isSafeListRowNavigationUrl(href: string): boolean {
  const t = href.trim();
  if (t.length === 0) return false;
  const lower = t.toLowerCase();
  if (lower.indexOf('javascript:') === 0 || lower.indexOf('data:') === 0) return false;
  if (t.charAt(0) === '/') return true;
  if (lower.indexOf('https://') === 0 || lower.indexOf('http://') === 0) return true;
  return false;
}
