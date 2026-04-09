import { fileFromServerRelativePath } from '@pnp/sp/files';
import type { IFolder } from '@pnp/sp/folders/types';

import { getSP } from '../../../../services/core/sp';
import type { IAttachmentLibraryFolderTreeNode, IFormManagerConfig } from '../config/types/formManager';
import { findUploadTargetId } from './attachmentFolderTree';

const PLACEHOLDER = /\{\{\s*([^}]+?)\s*\}\}/g;

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

function formatFieldValueForFolderToken(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && v !== null && 'Title' in v) {
    const t = (v as Record<string, unknown>).Title;
    return typeof t === 'string' ? t : String(t ?? '');
  }
  if (typeof v === 'object' && v !== null && 'LookupValue' in v) {
    return String((v as Record<string, unknown>).LookupValue ?? '');
  }
  return String(v);
}

export function resolveAttachmentFolderSegmentTemplate(
  template: string,
  itemId: number,
  itemFieldValues: Record<string, unknown>
): string {
  return template.replace(PLACEHOLDER, (_: string, rawKey: string) => {
    const key = rawKey.trim();
    if (/^itemid$/i.test(key)) return String(itemId);
    return formatFieldValueForFolderToken(itemFieldValues[key]);
  });
}

export function sanitizeSharePointFolderLeafName(name: string): string {
  let s = name
    .replace(/[\\/:*?"<>|#%]/g, ' ')
    .split('')
    .filter((ch) => ch.charCodeAt(0) >= 32)
    .join('')
    .replace(/\s+/g, ' ')
    .trim();
  if (s.length > 120) s = s.slice(0, 120).trim();
  return s;
}

async function ensureChildFolder(parent: IFolder, resolvedSegment: string): Promise<IFolder> {
  const name = sanitizeSharePointFolderLeafName(resolvedSegment);
  if (!name) {
    throw new Error('Um nível da pasta ficou vazio após resolver modelos ou sanitizar o nome.');
  }
  try {
    return await parent.addSubFolderUsingPath(name);
  } catch {
    return parent.folders.getByUrl(name);
  }
}

async function ensureFolderChainFromListRoot(list: { rootFolder: IFolder }, templates: string[]): Promise<IFolder> {
  let folder: IFolder = list.rootFolder;
  for (let i = 0; i < templates.length; i++) {
    folder = await ensureChildFolder(folder, templates[i]);
  }
  return folder;
}

async function ensureFolderTreeUnderParent(
  parent: IFolder,
  nodes: IAttachmentLibraryFolderTreeNode[],
  itemId: number,
  values: Record<string, unknown>,
  folderByNodeId: Map<string, IFolder>
): Promise<void> {
  for (let i = 0; i < nodes.length; i++) {
    const node = nodes[i];
    const resolved = resolveAttachmentFolderSegmentTemplate(node.nameTemplate, itemId, values);
    const childFolder = await ensureChildFolder(parent, resolved);
    folderByNodeId.set(node.id, childFolder);
    if (node.children?.length) {
      await ensureFolderTreeUnderParent(childFolder, node.children, itemId, values, folderByNodeId);
    }
  }
}

export interface IUploadToAttachmentLibraryOptions {
  folderPathSegments?: string[];
  folderTree?: IAttachmentLibraryFolderTreeNode[];
  itemFieldValues?: Record<string, unknown>;
}

export async function uploadFilesToAttachmentLibrary(
  libraryTitle: string,
  lookupFieldInternalName: string,
  mainItemId: number,
  files: File[],
  options?: IUploadToAttachmentLibraryOptions
): Promise<void> {
  if (!files.length) return;
  const sp = getSP();
  const list = sp.web.lists.getByTitle(libraryTitle.trim());
  const lookupKey = `${lookupFieldInternalName.trim()}Id`;
  const values = options?.itemFieldValues ?? {};
  const idFolder = sanitizeSharePointFolderLeafName(String(mainItemId));
  if (!idFolder) {
    throw new Error('ID do item inválido para nome da pasta.');
  }
  const root = list.rootFolder as IFolder;
  let uploadFolder: IFolder;

  if (options?.folderTree?.length) {
    const idFolderHandle = await ensureChildFolder(root, idFolder);
    const folderByNodeId = new Map<string, IFolder>();
    await ensureFolderTreeUnderParent(idFolderHandle, options.folderTree, mainItemId, values, folderByNodeId);
    const targetId = findUploadTargetId(options.folderTree);
    const target = targetId ? folderByNodeId.get(targetId) : undefined;
    uploadFolder = target ?? idFolderHandle;
  } else {
    const rawSeg = options?.folderPathSegments?.filter((s) => typeof s === 'string' && s.trim()) ?? [];
    const resolvedSub: string[] = [];
    for (let i = 0; i < rawSeg.length; i++) {
      resolvedSub.push(resolveAttachmentFolderSegmentTemplate(rawSeg[i].trim(), mainItemId, values));
    }
    const subFolders: string[] = [];
    for (let j = 0; j < resolvedSub.length; j++) {
      const leaf = sanitizeSharePointFolderLeafName(resolvedSub[j]);
      if (leaf) subFolders.push(leaf);
    }
    const fullChain = [idFolder, ...subFolders];
    uploadFolder = await ensureFolderChainFromListRoot({ rootFolder: root }, fullChain);
  }
  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    const body = await f.arrayBuffer();
    const fileInfo = await uploadFolder.files.addUsingPath(f.name, body, {
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
      [lookupKey]: mainItemId,
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
