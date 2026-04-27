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
