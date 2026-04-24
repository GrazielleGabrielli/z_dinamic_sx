import { getSP } from '../../../../services/core/sp';
import type {
  IFormLinkedChildFormConfig,
  IFormManagerConfig,
  IFormManagerPermissionBreakConfig,
  IFormPermissionBreakAssignment,
  IFormManagerPermissionBreakTargets,
} from '../config/types/formManager';
import { isFormAttachmentLibraryRuntime } from './formAttachmentLibrary';
import { resolveLinkedChildAttachmentRuntime } from './linkedChildAttachmentRuntime';
import type { ILinkedChildRowState } from './formLinkedChildSync';

function listRef(sp: ReturnType<typeof getSP>, titleOrId: string) {
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

const ROLE_NAME_ALIASES: Record<string, string[]> = {
  Read: ['Read', 'Leitura'],
  Contribute: ['Contribute', 'Contribuir'],
  Edit: ['Edit', 'Editar'],
  'Full Control': ['Full Control', 'Controlo total', 'Control total'],
};

async function resolveRoleDefinitionId(roleName: string): Promise<{ id: number; name: string } | undefined> {
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

async function clearAllRoleAssignments(listTitleOrId: string, itemId: number): Promise<void> {
  const sp = getSP();
  const itemAny = listRef(sp, listTitleOrId).items.getById(itemId) as unknown as Record<string, unknown>;
  const raCol = itemAny.roleAssignments as
    | (() => Promise<{ Id: number }[]>) & { getById: (id: number) => { delete: () => Promise<unknown> } }
    | undefined;
  if (!raCol || typeof raCol !== 'function') return;
  let ras: { Id: number }[] = [];
  try {
    ras = await raCol();
  } catch {
    return;
  }
  for (let i = 0; i < ras.length; i++) {
    try {
      await raCol.getById(ras[i].Id).delete();
    } catch {
      /* ignore */
    }
  }
}

async function applyUniquePermissionsToListItem(
  listTitleOrId: string,
  itemId: number,
  pb: IFormManagerPermissionBreakConfig,
  principalRolePairs: { principalId: number; roleDefId: number }[],
  authorIdHint?: number
): Promise<void> {
  const sp = getSP();
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
  await clearAllRoleAssignments(listTitleOrId, itemId);
  const authorNum =
    typeof authorIdHint === 'number' && isFinite(authorIdHint) && authorIdHint > 0
      ? authorIdHint
      : typeof meta.AuthorId === 'number' && isFinite(meta.AuthorId)
        ? meta.AuthorId
        : Number(meta.AuthorId);
  const authorId = isFinite(authorNum) && authorNum > 0 ? authorNum : undefined;
  if (pb.retainAuthor !== false && authorId !== undefined) {
    const ar = await resolveRoleDefinitionId(pb.authorRoleDefinitionName ?? 'Contribuir');
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
  roleIdCache: Map<string, number>
): Promise<IApplyPermissionBreakPrincipalPair[]> {
  const out: IApplyPermissionBreakPrincipalPair[] = [];
  const seen = new Set<string>();
  const list = assignments ?? [];
  for (let i = 0; i < list.length; i++) {
    const a = list[i];
    let roleId = roleIdCache.get(a.roleDefinitionName);
    if (roleId === undefined) {
      const r = await resolveRoleDefinitionId(a.roleDefinitionName);
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
  pairs: IApplyPermissionBreakPrincipalPair[]
): Promise<void> {
  const sp = getSP();
  const fld = `${lookupFieldInternalName.trim()}Id`;
  const rows = await listRef(sp, libraryTitle)
    .items.filter(`${fld} eq ${lookupNumericId}`)
    .select('Id', 'AuthorId')();
  const arr = Array.isArray(rows) ? rows : [];
  for (let i = 0; i < arr.length; i++) {
    const r = arr[i] as Record<string, unknown>;
    const id = typeof r.Id === 'number' ? r.Id : Number(r.Id);
    if (!isFinite(id)) continue;
    const aid = typeof r.AuthorId === 'number' ? r.AuthorId : Number(r.AuthorId);
    await applyUniquePermissionsToListItem(
      libraryTitle,
      id,
      pb,
      pairs,
      isFinite(aid) && aid > 0 ? aid : undefined
    );
  }
}

export interface IApplyFormPermissionBreakInput {
  formManager: IFormManagerConfig;
  listTitle: string;
  mainItemId: number;
  mainValues: Record<string, unknown>;
  mainAuthorId?: number;
  linkedConfigsSorted: IFormLinkedChildFormConfig[];
  linkedRowsById: Record<string, ILinkedChildRowState[]>;
}

export async function applyFormManagerPermissionBreak(input: IApplyFormPermissionBreakInput): Promise<void> {
  const pb = input.formManager.permissionBreak;
  if (!pb?.enabled) return;
  const targets = pb.targets ?? { mainListItem: true };
  const roleIdCache = new Map<string, number>();

  if (targets.mainListItem !== false) {
    const pairs = await buildPrincipalRolePairs(
      pb.assignments,
      { mainValues: input.mainValues },
      roleIdCache
    );
    await applyUniquePermissionsToListItem(
      input.listTitle,
      input.mainItemId,
      pb,
      pairs,
      input.mainAuthorId
    );
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
        roleIdCache
      );
      const lt = cfg.listTitle.trim();
      if (!lt) continue;
      await applyUniquePermissionsToListItem(lt, sid, pb, pairs, undefined);
    }
  }

  if (targets.mainAttachmentLibraryFiles === true && isFormAttachmentLibraryRuntime(input.formManager)) {
    const lib = input.formManager.attachmentLibrary!;
    const lt = lib.libraryTitle?.trim() ?? '';
    const lk = lib.sourceListLookupFieldInternalName?.trim() ?? '';
    if (lt && lk) {
      const pairs = await buildPrincipalRolePairs(pb.assignments, { mainValues: input.mainValues }, roleIdCache);
      await applyToLibraryItemsByLookup(lt, lk, input.mainItemId, pb, pairs);
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
          roleIdCache
        );
        await applyToLibraryItemsByLookup(lt, lk, sid, pb, pairs);
      }
    }
  }
}
