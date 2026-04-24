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

export const getSP = (context?: WebPartContext): SPFI => {
  if (context) {
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};
