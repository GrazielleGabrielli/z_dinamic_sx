import { fileFromServerRelativePath } from '@pnp/sp/files';
import type { IFolder } from '@pnp/sp/folders/types';

import { getSP } from '../../../../services/core/sp';
import type { IAttachmentLibraryFolderTreeNode, IFormManagerConfig } from '../config/types/formManager';
import { findUploadTargetId, setUploadTargetById, treeHasPerStepFolderUploaders } from './attachmentFolderTree';

const PLACEHOLDER = /\{\{\s*([^}]+?)\s*\}\}/g;

export function isFormAttachmentLibraryRuntime(formManager: IFormManagerConfig): boolean {
  const lib = formManager.attachmentLibrary;
  return (
    formManager.attachmentStorageKind === 'documentLibrary' &&
    !!lib?.libraryTitle?.trim() &&
    !!lib?.sourceListLookupFieldInternalName?.trim()
  );
}

export interface IAttachmentLibraryFileRow {
  fileName: string;
  fileUrl: string;
  /** Caminho server-relative do ficheiro (para filtrar por pasta na biblioteca). */
  fileRef: string;
}

/** Segmentos de pasta sob a pasta com nome = ID do item (sem o nome do ficheiro). */
export function parseFolderSegmentsUnderItemFolder(fileRef: string, itemId: number): string[] {
  const idFolder = sanitizeSharePointFolderLeafName(String(itemId));
  const normalized = fileRef.replace(/\\/g, '/');
  const parts = normalized.split('/').filter(Boolean);
  const idx = parts.findIndex((p) => p === idFolder);
  if (idx < 0) return [];
  const after = parts.slice(idx + 1);
  if (after.length <= 1) return [];
  return after.slice(0, -1);
}

/**
 * Caminho de pastas (apenas nomes resolvidos) desde a pasta com nome = ID do item até ao nó indicado.
 * URL absoluta: raiz da biblioteca + `/` + id + `/` + segmentos.
 */
export function buildAttachmentFolderAbsoluteUrl(opts: {
  libraryRootServerRelativeUrl: string;
  itemId: number | undefined;
  folderTree: IAttachmentLibraryFolderTreeNode[] | undefined;
  folderNodeId: string;
  itemFieldValues: Record<string, unknown>;
}): string | undefined {
  const root = opts.libraryRootServerRelativeUrl.trim();
  const id = opts.itemId;
  if (!root || id === undefined || typeof id !== 'number' || !isFinite(id)) return undefined;
  const tree = opts.folderTree;
  if (!tree?.length) return undefined;
  const nodeId = opts.folderNodeId.trim();
  if (!nodeId) return undefined;
  const segments = getResolvedFolderSegmentsForNode(tree, nodeId, id, opts.itemFieldValues);
  if (segments === undefined) return undefined;
  const idFolder = sanitizeSharePointFolderLeafName(String(id));
  if (!idFolder) return undefined;
  const base = root.replace(/\\/g, '/').replace(/\/$/, '');
  const path = [base, idFolder, ...segments].join('/');
  const rel = path.startsWith('/') ? path : `/${path}`;
  if (/^https?:\/\//i.test(rel)) return rel;
  const origin = typeof window !== 'undefined' ? window.location.origin : '';
  return `${origin}${rel}`;
}

export function getResolvedFolderSegmentsForNode(
  nodes: IAttachmentLibraryFolderTreeNode[],
  targetNodeId: string,
  itemId: number,
  itemFieldValues: Record<string, unknown>
): string[] | undefined {
  function walk(
    ns: IAttachmentLibraryFolderTreeNode[],
    acc: string[]
  ): string[] | undefined {
    for (let i = 0; i < ns.length; i++) {
      const n = ns[i];
      const resolved = resolveAttachmentFolderSegmentTemplate(n.nameTemplate, itemId, itemFieldValues);
      const seg = sanitizeSharePointFolderLeafName(resolved);
      if (!seg) {
        if (n.id === targetNodeId) return undefined;
        if (n.children?.length) {
          const d = walk(n.children, acc);
          if (d) return d;
        }
        continue;
      }
      const nextAcc = acc.concat([seg]);
      if (n.id === targetNodeId) return nextAcc;
      if (n.children?.length) {
        const d = walk(n.children, nextAcc);
        if (d) return d;
      }
    }
    return undefined;
  }
  return walk(nodes, []);
}

export function libraryFileRowBelongsToFolderNode(
  fileRef: string,
  folderNodeId: string,
  folderTree: IAttachmentLibraryFolderTreeNode[] | undefined,
  itemId: number,
  itemFieldValues: Record<string, unknown>
): boolean {
  if (!folderTree?.length) return true;
  const expected = getResolvedFolderSegmentsForNode(folderTree, folderNodeId, itemId, itemFieldValues);
  if (expected === undefined) return false;
  const actual = parseFolderSegmentsUnderItemFolder(fileRef, itemId);
  if (expected.length !== actual.length) return false;
  for (let j = 0; j < expected.length; j++) {
    if (expected[j] !== actual[j]) return false;
  }
  return true;
}

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

function findAttachmentTreeNodeById(
  nodes: IAttachmentLibraryFolderTreeNode[] | undefined,
  id: string
): IAttachmentLibraryFolderTreeNode | undefined {
  if (!nodes?.length) return undefined;
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes[i];
    if (n.id === id) return n;
    const d = findAttachmentTreeNodeById(n.children, id);
    if (d) return d;
  }
  return undefined;
}

export function resolvedFolderDisplayLabelForTreeNode(
  tree: IAttachmentLibraryFolderTreeNode[],
  nodeId: string,
  itemId: number,
  itemFieldValues: Record<string, unknown>
): string {
  const node = findAttachmentTreeNodeById(tree, nodeId);
  if (!node) return nodeId.trim() || 'Pasta';
  const raw = resolveAttachmentFolderSegmentTemplate(node.nameTemplate, itemId, itemFieldValues);
  const leaf = sanitizeSharePointFolderLeafName(raw);
  return leaf || node.nameTemplate.trim() || node.id;
}

function resolvedTargetFolderLabelFromTree(
  folderTree: IAttachmentLibraryFolderTreeNode[],
  itemId: number,
  vals: Record<string, unknown>
): string {
  const tid = findUploadTargetId(folderTree);
  if (!tid) return 'Biblioteca';
  return resolvedFolderDisplayLabelForTreeNode(folderTree, tid, itemId, vals);
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
  /** Etiqueta da pasta já resolvida (ex.: ao gravar por nó na árvore). */
  folderDisplayLabel?: string;
  /** Antes de cada ficheiro ser enviado para a biblioteca. */
  onUploadFileStart?: (info: { folderLabel: string; fileName: string }) => void;
}

export async function uploadFilesToAttachmentLibraryByFolderNodes(
  libraryTitle: string,
  lookupFieldInternalName: string,
  mainItemId: number,
  filesByFolderNodeId: Record<string, File[]>,
  folderTree: IAttachmentLibraryFolderTreeNode[] | undefined,
  options?: Omit<IUploadToAttachmentLibraryOptions, 'folderTree' | 'folderPathSegments'>
): Promise<void> {
  const entries = Object.entries(filesByFolderNodeId).filter(([, fs]) => fs.length > 0);
  if (!entries.length) return;
  if (!folderTree?.length) {
    const all = entries.flatMap(([, fs]) => fs);
    await uploadFilesToAttachmentLibrary(libraryTitle, lookupFieldInternalName, mainItemId, all, options);
    return;
  }
  if (!treeHasPerStepFolderUploaders(folderTree)) {
    const all = entries.flatMap(([, fs]) => fs);
    await uploadFilesToAttachmentLibrary(libraryTitle, lookupFieldInternalName, mainItemId, all, {
      ...options,
      folderTree,
      folderPathSegments: undefined,
    });
    return;
  }
  const values = options?.itemFieldValues ?? {};
  for (let i = 0; i < entries.length; i++) {
    const nodeId = entries[i][0];
    const files = entries[i][1];
    const treeForTarget = setUploadTargetById(folderTree, nodeId);
    const folderDisplayLabel = folderTree?.length
      ? resolvedFolderDisplayLabelForTreeNode(folderTree, nodeId, mainItemId, values)
      : undefined;
    await uploadFilesToAttachmentLibrary(libraryTitle, lookupFieldInternalName, mainItemId, files, {
      ...options,
      folderTree: treeForTarget,
      folderPathSegments: undefined,
      folderDisplayLabel,
    });
  }
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
  const folderLabelForUi = ((): string => {
    const ov = options?.folderDisplayLabel?.trim();
    if (ov) return ov;
    if (options?.folderTree?.length) return resolvedTargetFolderLabelFromTree(options.folderTree, mainItemId, values);
    const rawSeg = options?.folderPathSegments?.filter((s) => typeof s === 'string' && s.trim()) ?? [];
    if (rawSeg.length) {
      const resolvedSub: string[] = [];
      for (let i = 0; i < rawSeg.length; i++) {
        resolvedSub.push(resolveAttachmentFolderSegmentTemplate(rawSeg[i].trim(), mainItemId, values));
      }
      const parts = resolvedSub.map((s) => sanitizeSharePointFolderLeafName(s)).filter(Boolean);
      return parts.length ? parts.join(' › ') : 'Biblioteca';
    }
    return 'Biblioteca';
  })();
  const onUploadFileStart = options?.onUploadFileStart;
  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    onUploadFileStart?.({ folderLabel: folderLabelForUi, fileName: f.name });
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
): Promise<IAttachmentLibraryFileRow[]> {
  const sp = getSP();
  const list = sp.web.lists.getByTitle(libraryTitle.trim());
  const fld = `${lookupFieldInternalName.trim()}Id`;
  const filter = `${fld} eq ${mainItemId}`;
  const raw = await list.items.filter(filter).select('FileLeafRef', 'FileRef').top(5000)();
  const rows = Array.isArray(raw) ? raw : [];
  const out: IAttachmentLibraryFileRow[] = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] as Record<string, unknown>;
    const fn = r.FileLeafRef;
    const name = typeof fn === 'string' && fn.trim() ? fn.trim() : '';
    if (!name) continue;
    const sr = r.FileRef ?? r.ServerRelativeUrl;
    const fileRef = typeof sr === 'string' && sr.trim() ? sr.trim() : '';
    let fileUrl = '';
    if (fileRef) {
      fileUrl = serverRelativeToAbsoluteUrl(fileRef);
    }
    out.push({ fileName: name, fileUrl, fileRef });
  }
  return out;
}
