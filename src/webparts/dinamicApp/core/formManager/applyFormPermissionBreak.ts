import { fileFromServerRelativePath } from '@pnp/sp/files';

import type { SPFI } from '@pnp/sp';
import { getSP, getSPForWeb } from '../../../../services/core/sp';
import type {
  IFormLinkedChildFormConfig,
  IFormManagerConfig,
  IFormManagerPermissionBreakConfig,
  IFormPermissionBreakAssignment,
  IFormManagerPermissionBreakTargets,
} from '../config/types/formManager';
import {
  collectLibraryFolderListItemIdsUnderItemFolder,
  isFormAttachmentLibraryRuntime,
} from './formAttachmentLibrary';
import { resolveLinkedChildAttachmentRuntime } from './linkedChildAttachmentRuntime';
import type { ILinkedChildRowState } from './formLinkedChildSync';

function listRef(sp: SPFI, titleOrId: string) {
  const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(titleOrId);
  return isGuid ? sp.web.lists.getById(titleOrId) : sp.web.lists.getByTitle(titleOrId);
}

function userFieldIds(v: unknown): number[] {
  if (v === null || v === undefined) return [];
  if (typeof v === 'number' && isFinite(v)) return v > 0 ? [v] : [];
  if (Array.isArray(v)) {
    const out: number[] = [];
    for (let i = 0; i < v.length; i++) {
      out.push(...userFieldIds(v[i]));
    }
    return out.filter((x, j, a) => a.indexOf(x) === j);
  }
  if (typeof v === 'object' && v !== null && 'Id' in v) {
    const id = (v as Record<string, unknown>).Id;
    if (typeof id === 'number' && isFinite(id) && id > 0) return [id];
  }
  return [];
}

function asOdataArray(raw: unknown): Record<string, unknown>[] {
  if (Array.isArray(raw)) return raw as Record<string, unknown>[];
  if (raw && typeof raw === 'object' && Array.isArray((raw as { value?: unknown[] }).value)) {
    return (raw as { value: Record<string, unknown>[] }).value;
  }
  return [];
}

const ROLE_NAME_ALIASES: Record<string, string[]> = {
  Read: ['Read', 'Leitura'],
  Contribute: ['Contribute', 'Contribuir', 'Colaboração', 'Colaboracao'],
  Edit: ['Edit', 'Editar', 'Edição', 'Ediçao', 'Edicao'],
  'Full Control': ['Full Control', 'Controlo total', 'Control total', 'Controle total'],
};

function normRoleLookupKey(s: string): string {
  return s
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

async function fetchWebRoleDefinitionLookup(sp: SPFI): Promise<Map<string, number>> {
  const map = new Map<string, number>();
  try {
    const raw = await (sp.web as unknown as { roleDefinitions: { select: (...f: string[]) => { (): Promise<unknown> } } })
      .roleDefinitions.select('Id', 'Name')();
    const rows = asOdataArray(raw);
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i] as Record<string, unknown>;
      const id = typeof r.Id === 'number' ? r.Id : Number(r.Id);
      const name = typeof r.Name === 'string' ? r.Name.trim() : '';
      if (!isFinite(id) || id <= 0 || !name) continue;
      map.set(normRoleLookupKey(name), id);
    }
  } catch {
    /* getByName + aliases must suffice */
  }
  return map;
}

async function resolveRoleDefinitionId(
  roleName: string,
  siteLookup?: Map<string, number>
): Promise<{ id: number; name: string } | undefined> {
  const sp = getSP();
  const trimmed = roleName.trim();
  if (!trimmed) return undefined;
  const tryNames = new Set<string>([trimmed]);
  for (const [, aliases] of Object.entries(ROLE_NAME_ALIASES)) {
    if (aliases.some((a) => a.toLowerCase() === trimmed.toLowerCase())) {
      for (const a of aliases) tryNames.add(a);
    }
  }
  for (const n of Array.from(tryNames)) {
    try {
      const def = await sp.web.roleDefinitions.getByName(n).select('Id', 'Name')();
      const id = typeof def.Id === 'number' ? def.Id : Number(def.Id);
      if (isFinite(id) && id > 0) return { id, name: String(def.Name ?? n) };
    } catch {
      /* try next */
    }
  }
  if (siteLookup && siteLookup.size > 0) {
    for (const n of Array.from(tryNames)) {
      const id = siteLookup.get(normRoleLookupKey(n));
      if (id !== undefined && isFinite(id) && id > 0) return { id, name: n };
    }
    const id = siteLookup.get(normRoleLookupKey(trimmed));
    if (id !== undefined && isFinite(id) && id > 0) return { id, name: trimmed };
  }
  return undefined;
}

async function ensurePrincipalUserId(pickerKey: string): Promise<number | undefined> {
  const sp = getSP();
  try {
    const r = await sp.web.ensureUser(pickerKey);
    const o = r as { Id?: number; data?: { Id?: number } };
    const id = o.Id ?? o.data?.Id;
    if (typeof id === 'number' && isFinite(id) && id > 0) return id;
  } catch {
    return undefined;
  }
  return undefined;
}

function roleDefinitionBindingIds(ra: Record<string, unknown>): number[] {
  const raw = ra.RoleDefinitionBindings as Record<string, unknown> | undefined;
  if (!raw || typeof raw !== 'object') return [];
  const results = raw.results as unknown[] | undefined;
  const arr = Array.isArray(results) ? results : Array.isArray(raw) ? (raw as unknown[]) : [];
  const out: number[] = [];
  for (let i = 0; i < arr.length; i++) {
    const b = arr[i] as Record<string, unknown> | undefined;
    if (!b) continue;
    const id = typeof b.Id === 'number' ? b.Id : Number(b.Id);
    if (isFinite(id) && id > 0) out.push(id);
  }
  return out;
}

async function clearAllRoleAssignments(listTitleOrId: string, itemId: number, sp: SPFI = getSP()): Promise<void> {
  const item = listRef(sp, listTitleOrId).items.getById(itemId) as unknown as {
    roleAssignments: {
      expand: (s: string) => { (): Promise<Record<string, unknown>[]> };
      remove: (principalId: number, roleDefId: number) => Promise<unknown>;
    };
  };
  let ras: Record<string, unknown>[] = [];
  try {
    ras = asOdataArray(await item.roleAssignments.expand('RoleDefinitionBindings')());
  } catch {
    try {
      ras = asOdataArray(await (item.roleAssignments as unknown as { (): Promise<unknown> })());
    } catch {
      return;
    }
  }
  for (let i = 0; i < ras.length; i++) {
    const ra = ras[i];
    const pid = typeof ra.PrincipalId === 'number' ? ra.PrincipalId : Number(ra.PrincipalId);
    if (!isFinite(pid)) continue;
    const roleIds = roleDefinitionBindingIds(ra);
    if (roleIds.length === 0) continue;
    for (let j = 0; j < roleIds.length; j++) {
      try {
        await item.roleAssignments.remove(pid, roleIds[j]);
      } catch {
        /* ignore */
      }
    }
  }
}

async function applyUniquePermissionsToListItem(
  listTitleOrId: string,
  itemId: number,
  pb: IFormManagerPermissionBreakConfig,
  principalRolePairs: { principalId: number; roleDefId: number }[],
  authorIdHint?: number,
  siteRoleLookup?: Map<string, number>,
  sp: SPFI = getSP()
): Promise<void> {
  const itemAny = listRef(sp, listTitleOrId).items.getById(itemId) as unknown as {
    select: (...f: string[]) => { (): Promise<{ HasUniqueRoleAssignments?: boolean; AuthorId?: number }> };
    breakRoleInheritance: (copy: boolean, clearSubscopes?: boolean) => Promise<unknown>;
    roleAssignments: { add: (principalId: number, roleDefId: number) => Promise<unknown> };
  };
  const meta = await itemAny.select('HasUniqueRoleAssignments', 'AuthorId')();
  const copy = pb.copyInheritedAssignments === true;
  if (meta.HasUniqueRoleAssignments !== true) {
    await itemAny.breakRoleInheritance(copy, false);
  }
  await clearAllRoleAssignments(listTitleOrId, itemId, sp);
  const authorNum =
    typeof authorIdHint === 'number' && isFinite(authorIdHint) && authorIdHint > 0
      ? authorIdHint
      : typeof meta.AuthorId === 'number' && isFinite(meta.AuthorId)
        ? meta.AuthorId
        : Number(meta.AuthorId);
  const authorId = isFinite(authorNum) && authorNum > 0 ? authorNum : undefined;
  if (pb.retainAuthor !== false && authorId !== undefined) {
    const ar = await resolveRoleDefinitionId(pb.authorRoleDefinitionName ?? 'Contribuir', siteRoleLookup);
    if (ar) {
      try {
        await itemAny.roleAssignments.add(authorId, ar.id);
      } catch {
        /* ignore */
      }
    }
  }
  for (let j = 0; j < principalRolePairs.length; j++) {
    const p = principalRolePairs[j];
    try {
      await itemAny.roleAssignments.add(p.principalId, p.roleDefId);
    } catch {
      /* ignore duplicate */
    }
  }
}

async function applyUniquePermissionsToListItemAttachments(
  listTitleOrId: string,
  itemId: number,
  pb: IFormManagerPermissionBreakConfig,
  pairs: IApplyPermissionBreakPrincipalPair[],
  onProgress?: (detail: string) => void,
  siteRoleLookup?: Map<string, number>,
  sp: SPFI = getSP()
): Promise<void> {
  const item = listRef(sp, listTitleOrId).items.getById(itemId) as unknown as {
    attachmentFiles: { (): Promise<unknown> };
  };
  let raw: unknown;
  try {
    raw = await item.attachmentFiles();
  } catch {
    return;
  }
  const atts = asOdataArray(raw) as { ServerRelativeUrl?: string; FileName?: string }[];
  for (let i = 0; i < atts.length; i++) {
    const url = (atts[i].ServerRelativeUrl ?? '').trim();
    if (!url) continue;
    const fn = (atts[i].FileName ?? '').trim() || url.split('/').pop() || `#${i + 1}`;
    onProgress?.(`Quebra · anexo · ${fn}`);
    try {
      const file = fileFromServerRelativePath(sp.web, url);
      const attItem = await file.getItem<{ Id?: number; AuthorId?: number }>('Id', 'AuthorId');
      const aid = typeof attItem.Id === 'number' ? attItem.Id : Number(attItem.Id);
      if (!isFinite(aid) || aid <= 0) continue;
      const auth = typeof attItem.AuthorId === 'number' ? attItem.AuthorId : Number(attItem.AuthorId);
      await applyUniquePermissionsToListItem(
        listTitleOrId,
        aid,
        pb,
        pairs,
        isFinite(auth) && auth > 0 ? auth : undefined,
        siteRoleLookup,
        sp
      );
    } catch {
      /* */
    }
  }
}

export interface IApplyPermissionBreakPrincipalPair {
  principalId: number;
  roleDefId: number;
}

async function buildPrincipalRolePairs(
  assignments: IFormPermissionBreakAssignment[] | undefined,
  ctx: {
    mainValues: Record<string, unknown>;
    rowValues?: Record<string, unknown>;
    linkedFormIdForRow?: string;
  },
  roleIdCache: Map<string, number>,
  siteRoleLookup?: Map<string, number>
): Promise<IApplyPermissionBreakPrincipalPair[]> {
  const out: IApplyPermissionBreakPrincipalPair[] = [];
  const seen = new Set<string>();
  const list = assignments ?? [];
  for (let i = 0; i < list.length; i++) {
    const a = list[i];
    let roleId = roleIdCache.get(a.roleDefinitionName);
    if (roleId === undefined) {
      const r = await resolveRoleDefinitionId(a.roleDefinitionName, siteRoleLookup);
      if (!r) continue;
      roleId = r.id;
      roleIdCache.set(a.roleDefinitionName, roleId);
    }
    if (a.kind === 'siteGroup') {
      const gid = a.siteGroupId;
      if (gid === undefined || gid <= 0) continue;
      const key = `g:${gid}:${roleId}`;
      if (seen.has(key)) continue;
      seen.add(key);
      out.push({ principalId: gid, roleDefId: roleId });
    } else if (a.kind === 'user') {
      const pk = (a.userPickerKey ?? '').trim();
      if (!pk) continue;
      const uid = await ensurePrincipalUserId(pk);
      if (uid === undefined) continue;
      const key = `u:${uid}:${roleId}`;
      if (seen.has(key)) continue;
      seen.add(key);
      out.push({ principalId: uid, roleDefId: roleId });
    } else if (a.kind === 'field') {
      const fn = (a.fieldInternalName ?? '').trim();
      if (!fn) continue;
      const scope = a.fieldScope === 'linked' ? 'linked' : 'main';
      if (scope === 'linked') {
        if (!ctx.linkedFormIdForRow || a.linkedFormId !== ctx.linkedFormIdForRow) continue;
        const ids = userFieldIds(ctx.rowValues?.[fn]);
        for (let k = 0; k < ids.length; k++) {
          const key = `u:${ids[k]}:${roleId}`;
          if (seen.has(key)) continue;
          seen.add(key);
          out.push({ principalId: ids[k], roleDefId: roleId });
        }
      } else {
        const ids = userFieldIds(ctx.mainValues[fn]);
        for (let k = 0; k < ids.length; k++) {
          const key = `u:${ids[k]}:${roleId}`;
          if (seen.has(key)) continue;
          seen.add(key);
          out.push({ principalId: ids[k], roleDefId: roleId });
        }
      }
    }
  }
  return out;
}

function linkedFormIncluded(targets: IFormManagerPermissionBreakTargets | undefined, formId: string): boolean {
  const ids = targets?.linkedChildFormIds;
  if (Array.isArray(ids) && ids.length === 0) return false;
  if (!ids || ids.length === 0) return true;
  return ids.indexOf(formId) >= 0;
}

async function applyToLibraryItemsByLookup(
  libraryTitle: string,
  lookupFieldInternalName: string,
  lookupNumericId: number,
  pb: IFormManagerPermissionBreakConfig,
  pairs: IApplyPermissionBreakPrincipalPair[],
  onProgress?: (detail: string) => void,
  siteRoleLookup?: Map<string, number>,
  sp: SPFI = getSP()
): Promise<void> {
  const libLabel = libraryTitle.trim();
  onProgress?.(`Quebra · biblioteca «${libLabel}» · item-ligação ${lookupNumericId}`);
  const folderRefs = await collectLibraryFolderListItemIdsUnderItemFolder(libraryTitle, lookupNumericId);
  for (let fi = 0; fi < folderRefs.length; fi++) {
    const fr = folderRefs[fi];
    onProgress?.(`Quebra · biblioteca «${libLabel}» · pasta #${fr.id}`);
    await applyUniquePermissionsToListItem(
      libraryTitle,
      fr.id,
      pb,
      pairs,
      fr.authorId,
      siteRoleLookup,
      sp
    );
  }
  const fld = `${lookupFieldInternalName.trim()}Id`;
  const rows = await listRef(sp, libraryTitle)
    .items.filter(`${fld} eq ${lookupNumericId}`)
    .select('Id', 'AuthorId', 'FileLeafRef')
    .top(5000)();
  const arr = asOdataArray(rows);
  for (let i = 0; i < arr.length; i++) {
    const r = arr[i] as Record<string, unknown>;
    const id = typeof r.Id === 'number' ? r.Id : Number(r.Id);
    if (!isFinite(id)) continue;
    const leaf = typeof r.FileLeafRef === 'string' ? r.FileLeafRef.trim() : '';
    onProgress?.(
      leaf
        ? `Quebra · biblioteca «${libLabel}» · ficheiro ${leaf}`
        : `Quebra · biblioteca «${libLabel}» · ficheiro #${id}`
    );
    const aid = typeof r.AuthorId === 'number' ? r.AuthorId : Number(r.AuthorId);
    await applyUniquePermissionsToListItem(
      libraryTitle,
      id,
      pb,
      pairs,
      isFinite(aid) && aid > 0 ? aid : undefined,
      siteRoleLookup,
      sp
    );
  }
}

export interface IApplyFormPermissionBreakInput {
  formManager: IFormManagerConfig;
  listTitle: string;
  mainListWebServerRelativeUrl?: string;
  mainItemId: number;
  mainValues: Record<string, unknown>;
  mainAuthorId?: number;
  linkedConfigsSorted: IFormLinkedChildFormConfig[];
  linkedRowsById: Record<string, ILinkedChildRowState[]>;
  /** Detalhe na timeline (botão criar/atualizar), no estilo do upload (pasta · ficheiro). */
  onProgress?: (detail: string) => void;
}

export async function applyFormManagerPermissionBreak(input: IApplyFormPermissionBreakInput): Promise<void> {
  const pb = input.formManager.permissionBreak;
  if (!pb?.enabled) return;
  const targets = pb.targets ?? { mainListItem: true };
  const roleIdCache = new Map<string, number>();
  const prog = input.onProgress;
  const spDefault = getSP();
  const spMain = getSPForWeb(input.mainListWebServerRelativeUrl);
  const siteRoleLookup = await fetchWebRoleDefinitionLookup(spMain);

  if (targets.mainListItem !== false) {
    const pairs = await buildPrincipalRolePairs(
      pb.assignments,
      { mainValues: input.mainValues },
      roleIdCache,
      siteRoleLookup
    );
    prog?.(`Quebra · lista «${input.listTitle}» · item ${input.mainItemId}`);
    await applyUniquePermissionsToListItem(
      input.listTitle,
      input.mainItemId,
      pb,
      pairs,
      input.mainAuthorId,
      siteRoleLookup,
      spMain
    );
    if (input.formManager.attachmentStorageKind === 'itemAttachments') {
      await applyUniquePermissionsToListItemAttachments(
        input.listTitle,
        input.mainItemId,
        pb,
        pairs,
        prog,
        siteRoleLookup,
        spMain
      );
    }
  }

  for (let ci = 0; ci < input.linkedConfigsSorted.length; ci++) {
    const cfg = input.linkedConfigsSorted[ci];
    if (!linkedFormIncluded(targets, cfg.id)) continue;
    const rows = input.linkedRowsById[cfg.id] ?? [];
    for (let ri = 0; ri < rows.length; ri++) {
      const row = rows[ri];
      const sid = row.sharePointId;
      if (sid === undefined || !isFinite(sid)) continue;
      const pairs = await buildPrincipalRolePairs(
        pb.assignments,
        {
          mainValues: input.mainValues,
          rowValues: row.values,
          linkedFormIdForRow: cfg.id,
        },
        roleIdCache,
        siteRoleLookup
      );
      const lt = cfg.listTitle.trim();
      if (!lt) continue;
      prog?.(`Quebra · vinculada «${lt}» · item ${sid}`);
      await applyUniquePermissionsToListItem(lt, sid, pb, pairs, undefined, siteRoleLookup, spDefault);
      if (cfg.childAttachmentStorageKind === 'itemAttachments') {
        await applyUniquePermissionsToListItemAttachments(lt, sid, pb, pairs, prog, siteRoleLookup, spDefault);
      }
    }
  }

  if (targets.mainAttachmentLibraryFiles === true && isFormAttachmentLibraryRuntime(input.formManager)) {
    const lib = input.formManager.attachmentLibrary!;
    const lt = lib.libraryTitle?.trim() ?? '';
    const lk = lib.sourceListLookupFieldInternalName?.trim() ?? '';
    if (lt && lk) {
      const pairs = await buildPrincipalRolePairs(
        pb.assignments,
        { mainValues: input.mainValues },
        roleIdCache,
        siteRoleLookup
      );
      await applyToLibraryItemsByLookup(lt, lk, input.mainItemId, pb, pairs, prog, siteRoleLookup, spMain);
    }
  }

  if (targets.linkedAttachmentLibraryFilesByFormId?.length) {
    const allow = new Set(targets.linkedAttachmentLibraryFilesByFormId);
    for (let ci = 0; ci < input.linkedConfigsSorted.length; ci++) {
      const cfg = input.linkedConfigsSorted[ci];
      if (!allow.has(cfg.id)) continue;
      const resolved = resolveLinkedChildAttachmentRuntime(cfg, input.formManager);
      if (resolved.kind !== 'documentLibrary') continue;
      const lt = resolved.libraryTitle.trim();
      const lk = resolved.lookupFieldInternalName.trim();
      if (!lt || !lk) continue;
      const rows = input.linkedRowsById[cfg.id] ?? [];
      for (let ri = 0; ri < rows.length; ri++) {
        const row = rows[ri];
        const sid = row.sharePointId;
        if (sid === undefined || !isFinite(sid)) continue;
        const pairs = await buildPrincipalRolePairs(
          pb.assignments,
          {
            mainValues: input.mainValues,
            rowValues: row.values,
            linkedFormIdForRow: cfg.id,
          },
          roleIdCache,
          siteRoleLookup
        );
        prog?.(`Quebra · biblioteca (linha filha) «${lt}» · ligação ${sid}`);
        await applyToLibraryItemsByLookup(lt, lk, sid, pb, pairs, prog, siteRoleLookup, spDefault);
      }
    }
  }
}
