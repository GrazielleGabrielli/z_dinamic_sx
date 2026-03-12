import type { IDynamicContext } from './types';

export interface IBuildDynamicContextParams {
  currentUser?: {
    id?: number;
    title?: string;
    name?: string;
    email?: string;
    loginName?: string;
    department?: string;
    jobTitle?: string;
  };
  site?: { title?: string; url?: string };
  list?: { title?: string };
  query?: Record<string, string>;
  now?: Date;
}

/**
 * Monta IDynamicContext a partir de parâmetros (ex.: WebPartContext, window.location.search).
 * Útil para chamar de um único lugar que tem acesso a context e URL.
 */
export function buildDynamicContext(params: IBuildDynamicContextParams): IDynamicContext {
  return {
    currentUser: params.currentUser,
    site: params.site,
    list: params.list,
    query: params.query,
    now: params.now,
  };
}

/**
 * Parseia query string (ex.: window.location.search) em Record<string, string>.
 * Ex.: "?status=Pendente&id=1" → { status: 'Pendente', id: '1' }
 */
export function parseQueryString(search: string): Record<string, string> {
  const out: Record<string, string> = {};
  if (!search || typeof search !== 'string') return out;
  const trimmed = search.trim();
  if (trimmed.charAt(0) === '?') {
    const rest = trimmed.slice(1);
    const pairs = rest.split('&');
    for (let i = 0; i < pairs.length; i++) {
      const eq = pairs[i].indexOf('=');
      if (eq === -1) {
        const key = decodeURIComponent(pairs[i].trim());
        if (key) out[key] = '';
      } else {
        const key = decodeURIComponent(pairs[i].slice(0, eq).trim());
        const val = decodeURIComponent(pairs[i].slice(eq + 1).trim());
        if (key) out[key] = val;
      }
    }
  }
  return out;
}
