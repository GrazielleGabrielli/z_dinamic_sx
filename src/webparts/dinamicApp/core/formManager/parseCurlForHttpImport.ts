import type { IFormButtonActionHttpRequest } from '../config/types/formManager';

const METHODS = new Set(['GET', 'POST', 'PUT', 'PATCH', 'DELETE']);

function normalizeLineContinuations(s: string): string {
  return s.replace(/\\\r?\n[ \t]*/g, ' ').trim();
}

function tokenizeCurlLine(s: string): string[] {
  const tokens: string[] = [];
  let i = 0;
  const n = s.length;
  while (i < n) {
    while (i < n && s[i] === ' ') i++;
    if (i >= n) break;
    const q = s[i];
    if (q === '"' || q === "'") {
      i++;
      let buf = '';
      while (i < n) {
        if (s[i] === '\\' && i + 1 < n) {
          buf += s[i + 1];
          i += 2;
          continue;
        }
        if (s[i] === q) {
          i++;
          break;
        }
        buf += s[i];
        i++;
      }
      tokens.push(buf);
    } else {
      let buf = '';
      while (i < n && s[i] !== ' ') {
        if (s[i] === '\\' && i + 1 < n) {
          buf += s[i + 1];
          i += 2;
          continue;
        }
        buf += s[i];
        i++;
      }
      tokens.push(buf);
    }
  }
  return tokens;
}

function splitUrlBaseAndQuery(fullUrl: string): { base: string; query: Array<{ key: string; value: string }> } {
  const out: Array<{ key: string; value: string }> = [];
  const qm = fullUrl.indexOf('?');
  if (qm === -1) return { base: fullUrl.trim(), query: out };
  const base = fullUrl.slice(0, qm).trim();
  const raw = fullUrl.slice(qm + 1);
  if (!raw.trim()) return { base, query: out };
  if (fullUrl.includes('{{')) {
    return { base: fullUrl.trim(), query: out };
  }
  const pairs = raw.split('&');
  for (let p = 0; p < pairs.length; p++) {
    const part = pairs[p];
    if (!part) continue;
    const eq = part.indexOf('=');
    if (eq === -1) {
      try {
        out.push({ key: decodeURIComponent(part.replace(/\+/g, ' ')), value: '' });
      } catch {
        out.push({ key: part, value: '' });
      }
    } else {
      const k = part.slice(0, eq);
      const v = part.slice(eq + 1);
      try {
        out.push({
          key: decodeURIComponent(k.replace(/\+/g, ' ')),
          value: decodeURIComponent(v.replace(/\+/g, ' ')),
        });
      } catch {
        out.push({ key: k, value: v });
      }
    }
  }
  return { base, query: out };
}

function addHeader(headers: Array<{ key: string; value: string }>, key: string, value: string): void {
  const k = key.trim();
  if (!k) return;
  headers.push({ key: k, value: value.trim() });
}

export type TCurlImportResult =
  | {
      ok: true;
      method: IFormButtonActionHttpRequest['method'];
      url: string;
      headers: Array<{ key: string; value: string }>;
      query: Array<{ key: string; value: string }>;
      body: string;
    }
  | { ok: false; message: string };

export function parseCurlForHttpImport(raw: string): TCurlImportResult {
  const line = normalizeLineContinuations(raw);
  if (!line) return { ok: false, message: 'Cole um comando cURL.' };

  const tokens = tokenizeCurlLine(line);
  if (!tokens.length) return { ok: false, message: 'Comando vazio.' };

  let ti = 0;
  if (tokens[ti].toLowerCase() === 'curl') ti++;

  let method: IFormButtonActionHttpRequest['method'] = 'GET';
  let url = '';
  const headers: Array<{ key: string; value: string }> = [];
  let body = '';
  let bodySeen = false;

  while (ti < tokens.length) {
    const t = tokens[ti];
    const tl = t.toLowerCase();

    if (tl === '-x' || tl === '--request') {
      ti++;
      if (ti < tokens.length) {
        const up = tokens[ti].toUpperCase();
        if (METHODS.has(up)) method = up as IFormButtonActionHttpRequest['method'];
      }
      ti++;
      continue;
    }

    if (tl === '-h' || tl === '--header') {
      ti++;
      if (ti < tokens.length) {
        const hv = tokens[ti];
        const colon = hv.indexOf(':');
        if (colon > -1) addHeader(headers, hv.slice(0, colon), hv.slice(colon + 1));
        else addHeader(headers, hv, '');
      }
      ti++;
      continue;
    }

    if (tl.startsWith('--header=')) {
      const hv = t.slice('--header='.length);
      const colon = hv.indexOf(':');
      if (colon > -1) addHeader(headers, hv.slice(0, colon), hv.slice(colon + 1));
      else addHeader(headers, hv, '');
      ti++;
      continue;
    }

    if (tl === '-b' || tl === '--cookie') {
      ti++;
      if (ti < tokens.length) addHeader(headers, 'Cookie', tokens[ti]);
      ti++;
      continue;
    }

    if (
      tl === '-d' ||
      tl === '--data' ||
      tl === '--data-raw' ||
      tl === '--data-binary' ||
      tl.startsWith('--data=') ||
      tl.startsWith('--data-raw=') ||
      tl.startsWith('--data-binary=')
    ) {
      let payload: string | undefined;
      if (tl.startsWith('--data=')) payload = t.slice('--data='.length);
      else if (tl.startsWith('--data-raw=')) payload = t.slice('--data-raw='.length);
      else if (tl.startsWith('--data-binary=')) payload = t.slice('--data-binary='.length);
      else {
        ti++;
        if (ti < tokens.length) payload = tokens[ti];
      }
      if (payload !== undefined) {
        if (payload.startsWith('@')) {
          return { ok: false, message: 'Ficheiros (@ficheiro) em --data não são suportados; cole o corpo em texto.' };
        }
        body = payload;
        bodySeen = true;
        if (method === 'GET') method = 'POST';
      }
      ti++;
      continue;
    }

    if (tl === '--url') {
      ti++;
      if (ti < tokens.length) url = tokens[ti];
      ti++;
      continue;
    }

    if (tl.startsWith('-') && tl !== '-') {
      if (tl === '-g' || tl === '--get') {
        ti++;
        method = 'GET';
        continue;
      }
      ti++;
      if (tl === '-u' || tl === '--user' || tl === '-e' || tl === '--referer' || tl === '--referrer') {
        if (ti < tokens.length) ti++;
      }
      continue;
    }

    if (!url && (/^https?:\/\//i.test(t) || /^\/\//.test(t))) {
      url = t.startsWith('//') ? `https:${t}` : t;
      ti++;
      continue;
    }

    ti++;
  }

  if (!url.trim()) return { ok: false, message: 'Não foi encontrado um URL no comando.' };

  const { base, query } = splitUrlBaseAndQuery(url);

  if (bodySeen && method === 'GET') method = 'POST';

  return {
    ok: true,
    method,
    url: base,
    headers: headers.length ? headers : [{ key: '', value: '' }],
    query: query.length ? query : [{ key: '', value: '' }],
    body,
  };
}
