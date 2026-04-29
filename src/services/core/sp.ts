import { spfi, SPFI, SPFx } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/folders';
import '@pnp/sp/files';
import '@pnp/sp/items';
import '@pnp/sp/attachments';
import '@pnp/sp/fields';
import '@pnp/sp/views';
import '@pnp/sp/site-users/web';
import '@pnp/sp/site-groups/web';
import '@pnp/sp/profiles';
import '@pnp/sp/search';
import '@pnp/sp/security/item';
import '@pnp/sp/security/web';

let _sp: SPFI;
let _ctx: WebPartContext | undefined;

export const getSP = (context?: WebPartContext): SPFI => {
  if (context) {
    _ctx = context;
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};

function normWebPath(s: string): string {
  const t = (s || '').trim().replace(/\/+$/, '') || '/';
  return t.startsWith('/') ? t : `/${t}`;
}

export function getSPForWeb(webServerRelativeUrl?: string | null): SPFI {
  if (!_ctx) {
    return _sp;
  }
  const cur = normWebPath(_ctx.pageContext.web.serverRelativeUrl || '/');
  const target = normWebPath((webServerRelativeUrl ?? '').trim() || cur);
  if (target === cur) {
    return _sp;
  }
  const abs = _ctx.pageContext.web.absoluteUrl;
  if (!abs.endsWith(cur)) {
    return _sp;
  }
  const prefix = abs.slice(0, abs.length - cur.length);
  const targetAbs = `${prefix}${target.startsWith('/') ? target : `/${target}`}`;
  return spfi(targetAbs).using(SPFx(_ctx));
}

/** Identificador GUID de lista SharePoint (lista ligada por Id). Aceita com ou sem chaves. */
export function isSharePointListGuid(titleOrId: string): boolean {
  return /^\{?[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}\}?$/i.test((titleOrId || '').trim());
}

/** Remove chaves opcionais do GUID retornado pelo SP REST ({guid} → guid). */
export function normalizeListGuid(titleOrId: string): string {
  return (titleOrId || '').trim().replace(/^\{|\}$/g, '');
}

function parentServerRelativeWebSegment(pathNorm: string): string | undefined {
  const t = (pathNorm || '').trim().replace(/\/+$/, '');
  if (t === '' || t === '/') return undefined;
  const i = t.lastIndexOf('/');
  if (i <= 0) return '/';
  const p = t.slice(0, i).replace(/\/+$/, '') || '/';
  return p;
}

/**
 * Webs a experimentar para localizar uma lista por GUID (site / subsites).
 * Lista ligada pode viver no web pai; não no subsite da lista principal.
 */
export function buildWebPathCandidatesForListByGuid(
  preferredWebServerRelativeUrl?: string | null
): (string | undefined)[] {
  const out: (string | undefined)[] = [];
  const seen = new Set<string | undefined>();
  const push = (v: string | undefined): void => {
    if (seen.has(v)) return;
    seen.add(v);
    out.push(v);
  };

  const prefTrim = (preferredWebServerRelativeUrl ?? '').trim();
  if (!prefTrim) {
    push(undefined);
    return out;
  }

  let cur: string | undefined = normWebPath(prefTrim);
  while (cur !== undefined) {
    push(cur);
    const next = parentServerRelativeWebSegment(cur);
    if (next === undefined) break;
    if (next === cur) break;
    cur = next;
  }
  push(undefined);
  return out;
}
