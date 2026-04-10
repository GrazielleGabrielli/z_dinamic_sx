import type {
  IAttachmentLibraryFolderTreeNode,
  IFormManagerAttachmentLibraryConfig,
  TFormManagerFormMode,
} from '../config/types/formManager';
import { sanitizeFolderNameTemplatePreservingPlaceholders } from './attachmentFolderNameTemplate';
import { sanitizeConditionNode } from './formConditionSanitize';

export const MAX_ATTACHMENT_FOLDER_TREE_NODES = 40;
export const MAX_ATTACHMENT_FOLDER_TREE_DEPTH = 12;

export function newAttachmentFolderNodeId(): string {
  return `fld_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 9)}`;
}

export function createEmptyFolderNode(nameTemplate = ''): IAttachmentLibraryFolderTreeNode {
  return {
    id: newAttachmentFolderNodeId(),
    nameTemplate,
    children: undefined,
  };
}

function stripLeadingRedundantItemIdTemplate(segments: string[]): string[] {
  const out = segments.slice();
  while (out.length > 0 && /^\{\{\s*ItemId\s*\}\}$/i.test(out[0].trim())) {
    out.shift();
  }
  return out;
}

/** Converte lista linear legada numa única cadeia (um ramo) com destino na última pasta. */
export function migrateFolderPathSegmentsToTree(segments: string[]): IAttachmentLibraryFolderTreeNode[] {
  const s = stripLeadingRedundantItemIdTemplate(
    segments.map((x) => String(x).trim()).filter(Boolean)
  );
  if (!s.length) return [];
  let leaf: IAttachmentLibraryFolderTreeNode = {
    id: newAttachmentFolderNodeId(),
    nameTemplate: s[s.length - 1],
    uploadTarget: true,
  };
  for (let i = s.length - 2; i >= 0; i--) {
    leaf = {
      id: newAttachmentFolderNodeId(),
      nameTemplate: s[i],
      children: [leaf],
    };
  }
  return [leaf];
}

export function countNodesInTree(nodes: IAttachmentLibraryFolderTreeNode[]): number {
  let n = 0;
  for (let i = 0; i < nodes.length; i++) {
    n += 1;
    if (nodes[i].children?.length) n += countNodesInTree(nodes[i].children as IAttachmentLibraryFolderTreeNode[]);
  }
  return n;
}

function treeMaxDepth(nodes: IAttachmentLibraryFolderTreeNode[]): number {
  if (!nodes.length) return 0;
  let max = 0;
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes[i];
    const sub = n.children?.length ? treeMaxDepth(n.children) : 0;
    const here = 1 + sub;
    if (here > max) max = here;
  }
  return max;
}

function findFirstUploadTargetPreorder(
  nodes: IAttachmentLibraryFolderTreeNode[]
): IAttachmentLibraryFolderTreeNode | undefined {
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes[i];
    if (n.uploadTarget) return n;
    if (n.children?.length) {
      const d = findFirstUploadTargetPreorder(n.children);
      if (d) return d;
    }
  }
  return undefined;
}

function clearUploadTargets(nodes: IAttachmentLibraryFolderTreeNode[]): IAttachmentLibraryFolderTreeNode[] {
  return nodes.map((n) => ({
    ...n,
    uploadTarget: false,
    children: n.children?.length ? clearUploadTargets(n.children) : undefined,
  }));
}

function firstLeafPreorder(nodes: IAttachmentLibraryFolderTreeNode[]): IAttachmentLibraryFolderTreeNode | undefined {
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes[i];
    if (!n.children?.length) return n;
    const d = firstLeafPreorder(n.children);
    if (d) return d;
  }
  return undefined;
}

function setUploadTargetOnId(
  nodes: IAttachmentLibraryFolderTreeNode[],
  targetId: string,
  on: boolean
): IAttachmentLibraryFolderTreeNode[] {
  return nodes.map((n) => {
    if (n.id === targetId) {
      return { ...n, uploadTarget: on };
    }
    return {
      ...n,
      children: n.children?.length ? setUploadTargetOnId(n.children, targetId, on) : undefined,
    };
  });
}

/** Garante um único uploadTarget; se nenhum, marca a primeira folha em pré-ordem. */
export function normalizeFolderTreeUploadTarget(nodes: IAttachmentLibraryFolderTreeNode[]): IAttachmentLibraryFolderTreeNode[] {
  if (!nodes.length) return [];
  const firstMarked = findFirstUploadTargetPreorder(nodes);
  let cleared = clearUploadTargets(nodes);
  if (firstMarked) {
    return setUploadTargetOnId(cleared, firstMarked.id, true);
  }
  const leaf = firstLeafPreorder(cleared);
  if (leaf) {
    return setUploadTargetOnId(cleared, leaf.id, true);
  }
  return cleared;
}

function sanitizeNode(
  raw: unknown,
  depth: number,
  nodeCount: { n: number }
): IAttachmentLibraryFolderTreeNode | undefined {
  if (!raw || typeof raw !== 'object' || depth > MAX_ATTACHMENT_FOLDER_TREE_DEPTH) return undefined;
  const o = raw as Record<string, unknown>;
  const id = typeof o.id === 'string' && o.id.trim() ? o.id.trim().slice(0, 80) : newAttachmentFolderNodeId();
  const nameTemplate =
    typeof o.nameTemplate === 'string'
      ? sanitizeFolderNameTemplatePreservingPlaceholders(String(o.nameTemplate)).slice(0, 200)
      : '';
  if (nodeCount.n >= MAX_ATTACHMENT_FOLDER_TREE_NODES) return undefined;
  nodeCount.n += 1;
  let children: IAttachmentLibraryFolderTreeNode[] | undefined;
  if (Array.isArray(o.children) && o.children.length) {
    const ch: IAttachmentLibraryFolderTreeNode[] = [];
    for (let i = 0; i < o.children.length; i++) {
      const c = sanitizeNode(o.children[i], depth + 1, nodeCount);
      if (c) ch.push(c);
    }
    if (ch.length) children = ch;
  }
  let showUploaderInStepIds: string[] | undefined;
  if (Array.isArray(o.showUploaderInStepIds)) {
    const ids = o.showUploaderInStepIds
      .map((x) => String(x).trim())
      .filter(Boolean)
      .slice(0, 40);
    if (ids.length) showUploaderInStepIds = [ids[0]];
  }
  const showUploaderWhen = sanitizeConditionNode(o.showUploaderWhen);
  let showUploaderGroupTitles: string[] | undefined;
  if (Array.isArray(o.showUploaderGroupTitles)) {
    const g = o.showUploaderGroupTitles
      .map((x) => String(x).trim())
      .filter(Boolean)
      .slice(0, 80);
    if (g.length) showUploaderGroupTitles = g;
  }
  let showUploaderModes: TFormManagerFormMode[] | undefined;
  if (Array.isArray(o.showUploaderModes)) {
    const m = o.showUploaderModes.filter(
      (x): x is TFormManagerFormMode => x === 'create' || x === 'edit' || x === 'view'
    );
    if (m.length) showUploaderModes = m.slice(0, 3);
  }
  const ATT_LIM_CAP = 500;
  let minAttachmentCount: number | undefined;
  let maxAttachmentCount: number | undefined;
  if (typeof o.minAttachmentCount === 'number' && Number.isFinite(o.minAttachmentCount)) {
    minAttachmentCount = Math.max(0, Math.min(ATT_LIM_CAP, Math.floor(o.minAttachmentCount)));
  }
  if (typeof o.maxAttachmentCount === 'number' && Number.isFinite(o.maxAttachmentCount)) {
    maxAttachmentCount = Math.max(0, Math.min(ATT_LIM_CAP, Math.floor(o.maxAttachmentCount)));
  }
  if (minAttachmentCount !== undefined && maxAttachmentCount !== undefined && maxAttachmentCount < minAttachmentCount) {
    maxAttachmentCount = minAttachmentCount;
  }
  return {
    id,
    nameTemplate,
    ...(o.uploadTarget === true ? { uploadTarget: true } : {}),
    ...(showUploaderInStepIds ? { showUploaderInStepIds } : {}),
    ...(showUploaderWhen ? { showUploaderWhen } : {}),
    ...(showUploaderGroupTitles ? { showUploaderGroupTitles } : {}),
    ...(showUploaderModes ? { showUploaderModes } : {}),
    ...(minAttachmentCount !== undefined && minAttachmentCount > 0 ? { minAttachmentCount } : {}),
    ...(maxAttachmentCount !== undefined ? { maxAttachmentCount } : {}),
    ...(children ? { children } : {}),
  };
}

export function sanitizeFolderTreeInput(raw: unknown): IAttachmentLibraryFolderTreeNode[] {
  if (!Array.isArray(raw)) return [];
  const nodeCount = { n: 0 };
  const out: IAttachmentLibraryFolderTreeNode[] = [];
  for (let i = 0; i < raw.length; i++) {
    const n = sanitizeNode(raw[i], 0, nodeCount);
    if (n) out.push(n);
  }
  return normalizeFolderTreeUploadTarget(out);
}

export function loadFolderTreeFromAttachmentLibrary(
  lib: IFormManagerAttachmentLibraryConfig | undefined
): IAttachmentLibraryFolderTreeNode[] {
  if (!lib) return [];
  if (lib.folderTree?.length) return sanitizeFolderTreeInput(lib.folderTree);
  if (lib.folderPathSegments?.length) return sanitizeFolderTreeInput(migrateFolderPathSegmentsToTree(lib.folderPathSegments));
  return [];
}

export function attachmentFolderNodePathLabel(
  tree: IAttachmentLibraryFolderTreeNode[],
  targetId: string
): string {
  function walk(
    ns: IAttachmentLibraryFolderTreeNode[],
    acc: string[]
  ): string | undefined {
    for (let i = 0; i < ns.length; i++) {
      const n = ns[i];
      const label = (n.nameTemplate ?? '').trim() || '(pasta)';
      const next = acc.concat([label]);
      if (n.id === targetId) return next.join(' / ');
      if (n.children?.length) {
        const d = walk(n.children, next);
        if (d) return d;
      }
    }
    return undefined;
  }
  return walk(tree, []) ?? targetId;
}

export function flattenFolderTreeNodes(nodes: IAttachmentLibraryFolderTreeNode[]): IAttachmentLibraryFolderTreeNode[] {
  const out: IAttachmentLibraryFolderTreeNode[] = [];
  function walk(ns: IAttachmentLibraryFolderTreeNode[]): void {
    for (let i = 0; i < ns.length; i++) {
      out.push(ns[i]);
      if (ns[i].children?.length) walk(ns[i].children as IAttachmentLibraryFolderTreeNode[]);
    }
  }
  walk(nodes);
  return out;
}

export const FOLDER_ATTACHMENT_LIMIT_ERROR_PREFIX = '_attf_';

/** Mantém só erros de pastas cujo uploader está associado à etapa `stepId`. */
export function filterFolderLimitErrorsToStep(
  folderLimitErrors: Record<string, string>,
  tree: IAttachmentLibraryFolderTreeNode[] | undefined,
  stepId: string | undefined
): Record<string, string> {
  if (!tree?.length || !stepId) return {};
  const byId = new Map(flattenFolderTreeNodes(tree).map((n) => [n.id, n]));
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(folderLimitErrors)) {
    if (k.indexOf(FOLDER_ATTACHMENT_LIMIT_ERROR_PREFIX) !== 0) continue;
    const nid = k.slice(FOLDER_ATTACHMENT_LIMIT_ERROR_PREFIX.length);
    const n = byId.get(nid);
    if (n?.showUploaderInStepIds?.includes(stepId)) out[k] = v;
  }
  return out;
}

export function collectFolderAttachmentLimitErrors(
  tree: IAttachmentLibraryFolderTreeNode[] | undefined,
  opts: {
    pendingByFolder: Record<string, File[]>;
    libraryCountByNodeId: (nodeId: string) => number;
    isFolderUploaderVisible: (n: IAttachmentLibraryFolderTreeNode) => boolean;
  }
): Record<string, string> {
  const out: Record<string, string> = {};
  if (!tree?.length) return out;
  const flat = flattenFolderTreeNodes(tree);
  for (let i = 0; i < flat.length; i++) {
    const n = flat[i];
    const minC = n.minAttachmentCount;
    const maxC = n.maxAttachmentCount;
    const hasMin = minC !== undefined && minC > 0;
    const hasMax = maxC !== undefined;
    if (!hasMin && !hasMax) continue;
    if (!opts.isFolderUploaderVisible(n)) continue;
    const lib = opts.libraryCountByNodeId(n.id);
    const pend = opts.pendingByFolder[n.id]?.length ?? 0;
    const t = lib + pend;
    const label = n.nameTemplate?.trim() || 'Pasta';
    const key = `${FOLDER_ATTACHMENT_LIMIT_ERROR_PREFIX}${n.id}`;
    if (hasMin && t < minC!) {
      out[key] = `«${label}»: mínimo ${minC} ficheiro(s) (atual: ${t}).`;
    } else if (hasMax && t > maxC!) {
      out[key] = `«${label}»: máximo ${maxC} ficheiro(s) (atual: ${t}).`;
    }
  }
  return out;
}

export function treeHasPerStepFolderUploaders(nodes: IAttachmentLibraryFolderTreeNode[]): boolean {
  const flat = flattenFolderTreeNodes(nodes);
  for (let i = 0; i < flat.length; i++) {
    if (flat[i].showUploaderInStepIds?.length) return true;
  }
  return false;
}

export function patchNodeShowUploaderStepIds(
  nodes: IAttachmentLibraryFolderTreeNode[],
  id: string,
  stepIds: string[]
): IAttachmentLibraryFolderTreeNode[] {
  return nodes.map((n) => {
    if (n.id === id) {
      if (stepIds.length) return { ...n, showUploaderInStepIds: [stepIds[0]] };
      const { showUploaderInStepIds: _drop, ...rest } = n;
      return rest;
    }
    return {
      ...n,
      children: n.children?.length ? patchNodeShowUploaderStepIds(n.children, id, stepIds) : undefined,
    };
  });
}

export function updateAttachmentFolderNode(
  nodes: IAttachmentLibraryFolderTreeNode[],
  id: string,
  updater: (n: IAttachmentLibraryFolderTreeNode) => IAttachmentLibraryFolderTreeNode
): IAttachmentLibraryFolderTreeNode[] {
  return nodes.map((n) => {
    if (n.id === id) return updater(n);
    if (n.children?.length) {
      return { ...n, children: updateAttachmentFolderNode(n.children, id, updater) };
    }
    return n;
  });
}

export function findUploadTargetId(nodes: IAttachmentLibraryFolderTreeNode[]): string | undefined {
  for (let i = 0; i < nodes.length; i++) {
    if (nodes[i].uploadTarget) return nodes[i].id;
    if (nodes[i].children?.length) {
      const d = findUploadTargetId(nodes[i].children as IAttachmentLibraryFolderTreeNode[]);
      if (d) return d;
    }
  }
  return undefined;
}

export function addRootSibling(nodes: IAttachmentLibraryFolderTreeNode[]): IAttachmentLibraryFolderTreeNode[] {
  if (countNodesInTree(nodes) >= MAX_ATTACHMENT_FOLDER_TREE_NODES) return nodes;
  if (!nodes.length) {
    const n = createEmptyFolderNode('');
    n.uploadTarget = true;
    const out = [n];
    if (treeMaxDepth(out) > MAX_ATTACHMENT_FOLDER_TREE_DEPTH) return nodes;
    return out;
  }
  const out = nodes.concat([createEmptyFolderNode('')]);
  if (treeMaxDepth(out) > MAX_ATTACHMENT_FOLDER_TREE_DEPTH) return nodes;
  return normalizeFolderTreeUploadTarget(out);
}

export function addChild(
  nodes: IAttachmentLibraryFolderTreeNode[],
  parentId: string
): IAttachmentLibraryFolderTreeNode[] {
  if (countNodesInTree(nodes) >= MAX_ATTACHMENT_FOLDER_TREE_NODES) return nodes;
  const child = createEmptyFolderNode('');
  const next = nodes.map((n) => {
    if (n.id === parentId) {
      const ch = n.children?.slice() ?? [];
      return {
        ...n,
        children: ch.concat([child]),
      };
    }
    if (n.children?.length) {
      return { ...n, children: addChild(n.children, parentId) };
    }
    return n;
  });
  if (treeMaxDepth(next) > MAX_ATTACHMENT_FOLDER_TREE_DEPTH) return nodes;
  return next;
}

export function addSiblingAfter(
  nodes: IAttachmentLibraryFolderTreeNode[],
  afterId: string,
  isFolderTreeRootLevel = true
): IAttachmentLibraryFolderTreeNode[] {
  if (countNodesInTree(nodes) >= MAX_ATTACHMENT_FOLDER_TREE_NODES) return nodes;
  const idx = nodes.findIndex((n) => n.id === afterId);
  if (idx >= 0) {
    const next = nodes.slice();
    next.splice(idx + 1, 0, createEmptyFolderNode(''));
    if (treeMaxDepth(next) > MAX_ATTACHMENT_FOLDER_TREE_DEPTH) return nodes;
    return isFolderTreeRootLevel ? normalizeFolderTreeUploadTarget(next) : next;
  }
  const out = nodes.map((n) =>
    n.children?.length ? { ...n, children: addSiblingAfter(n.children, afterId, false) } : n
  );
  if (treeMaxDepth(out) > MAX_ATTACHMENT_FOLDER_TREE_DEPTH) return nodes;
  return out;
}

export function removeNodeById(
  nodes: IAttachmentLibraryFolderTreeNode[],
  id: string
): IAttachmentLibraryFolderTreeNode[] {
  const filtered = nodes.filter((n) => n.id !== id);
  if (filtered.length !== nodes.length) return normalizeFolderTreeUploadTarget(filtered);
  return nodes.map((n) => ({
    ...n,
    children: n.children?.length ? removeNodeById(n.children, id) : undefined,
  }));
}

export function patchNodeName(
  nodes: IAttachmentLibraryFolderTreeNode[],
  id: string,
  nameTemplate: string
): IAttachmentLibraryFolderTreeNode[] {
  return nodes.map((n) => {
    if (n.id === id) return { ...n, nameTemplate };
    return {
      ...n,
      children: n.children?.length ? patchNodeName(n.children, id, nameTemplate) : undefined,
    };
  });
}

export function setUploadTargetById(
  nodes: IAttachmentLibraryFolderTreeNode[],
  id: string
): IAttachmentLibraryFolderTreeNode[] {
  let next = clearUploadTargets(nodes);
  next = setUploadTargetOnId(next, id, true);
  return next;
}
