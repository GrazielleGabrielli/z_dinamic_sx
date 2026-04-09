import { fileFromServerRelativePath } from '@pnp/sp/files';

import { getSP } from '../../../../services/core/sp';
import type { IFormManagerConfig } from '../config/types/formManager';

export function isFormAttachmentLibraryRuntime(formManager: IFormManagerConfig): boolean {
  const lib = formManager.attachmentLibrary;
  return (
    formManager.attachmentStorageKind === 'documentLibrary' &&
    !!lib?.libraryTitle?.trim() &&
    !!lib?.sourceListLookupFieldInternalName?.trim()
  );
}

type IAttachmentRow = { fileName: string; fileUrl: string };

function serverRelativeToAbsoluteUrl(serverRelative: string): string {
  const path = serverRelative.trim();
  if (/^https?:\/\//i.test(path)) return path;
  return `${typeof window !== 'undefined' ? window.location.origin : ''}${
    path.startsWith('/') ? '' : '/'
  }${path}`;
}

export async function uploadFilesToAttachmentLibrary(
  libraryTitle: string,
  lookupFieldInternalName: string,
  mainListItemId: number,
  files: File[]
): Promise<void> {
  if (!files.length) return;
  const sp = getSP();
  const list = sp.web.lists.getByTitle(libraryTitle.trim());
  const lookupKey = `${lookupFieldInternalName.trim()}Id`;
  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    const body = await f.arrayBuffer();
    const fileInfo = await list.rootFolder.files.addUsingPath(f.name, body, {
      EnsureUniqueFileName: true,
    });
    const rel = (fileInfo as { ServerRelativeUrl?: string }).ServerRelativeUrl;
    if (!rel || !rel.trim()) {
      throw new Error('Upload sem ServerRelativeUrl');
    }
    const fileObj = fileFromServerRelativePath(sp.web, rel.trim());
    const item = await fileObj.getItem<{ Id?: number }>('Id');
    const libItemId = typeof item.Id === 'number' && isFinite(item.Id) ? item.Id : undefined;
    if (libItemId === undefined) {
      throw new Error('Upload sem Id do item na biblioteca');
    }
    await list.items.getById(libItemId).update({
      [lookupKey]: mainListItemId,
    });
  }
}

export async function loadLibraryAttachmentRowsForMainItem(
  libraryTitle: string,
  lookupFieldInternalName: string,
  mainItemId: number
): Promise<IAttachmentRow[]> {
  const sp = getSP();
  const list = sp.web.lists.getByTitle(libraryTitle.trim());
  const fld = `${lookupFieldInternalName.trim()}Id`;
  const filter = `${fld} eq ${mainItemId}`;
  const raw = await list.items.filter(filter).select('FileLeafRef', 'FileRef').top(5000)();
  const rows = Array.isArray(raw) ? raw : [];
  const out: IAttachmentRow[] = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] as Record<string, unknown>;
    const fn = r.FileLeafRef;
    const name = typeof fn === 'string' && fn.trim() ? fn.trim() : '';
    if (!name) continue;
    const sr = r.FileRef ?? r.ServerRelativeUrl;
    let fileUrl = '';
    if (typeof sr === 'string' && sr.trim()) {
      fileUrl = serverRelativeToAbsoluteUrl(sr.trim());
    }
    out.push({ fileName: name, fileUrl });
  }
  return out;
}
