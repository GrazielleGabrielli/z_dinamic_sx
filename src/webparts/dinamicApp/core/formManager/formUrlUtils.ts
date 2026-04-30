export function ensureAbsoluteSharePointUrl(raw: string): string {
  const s = raw.trim();
  if (!s) return '';
  if (/^https?:\/\//i.test(s)) return s;
  if (s.startsWith('//')) return s;
  if (typeof window !== 'undefined') {
    const path = s.startsWith('/') ? s : `/${s}`;
    return `${window.location.origin}${path}`;
  }
  return s.startsWith('/') ? s : `/${s}`;
}

export function parseUrlFieldValue(v: unknown): { Url: string; Description: string } {
  if (v === null || v === undefined) return { Url: '', Description: '' };
  if (typeof v === 'object' && v !== null && 'Url' in v) {
    const o = v as Record<string, unknown>;
    return {
      Url: ensureAbsoluteSharePointUrl(String(o.Url ?? '')),
      Description: String(o.Description ?? ''),
    };
  }
  const s = String(v);
  const comma = s.indexOf(',');
  if (comma !== -1) {
    return {
      Url: ensureAbsoluteSharePointUrl(s.slice(0, comma).trim()),
      Description: s.slice(comma + 1).trim(),
    };
  }
  return { Url: ensureAbsoluteSharePointUrl(s), Description: '' };
}
