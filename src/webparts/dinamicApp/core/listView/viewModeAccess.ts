import type { IListViewModeConfig } from '../config/types';

export function normWebPath(s: string): string {
  const t = (s || '').trim().replace(/\/+$/, '') || '/';
  return t.startsWith('/') ? t : `/${t}`;
}

export function resolveAccessWebKey(modeWeb: string | undefined | null, pageWeb: string): string {
  return normWebPath(modeWeb || pageWeb);
}

export function modeRequiresAccessCheck(mode: IListViewModeConfig): boolean {
  return mode.access !== undefined && mode.access !== null;
}

export function collectDistinctAccessWebKeys(modes: IListViewModeConfig[], pageWeb: string): string[] {
  const keys = new Set<string>();
  const pg = normWebPath(pageWeb);
  keys.add(pg);
  for (let i = 0; i < modes.length; i++) {
    const a = modes[i].access;
    if (!a || (a.allowedGroupIds?.length ?? 0) === 0) continue;
    const w = a.webServerRelativeUrl;
    keys.add(resolveAccessWebKey(w, pageWeb));
  }
  return Array.from(keys);
}

export function userCanUseViewMode(
  mode: IListViewModeConfig,
  currentUserId: number,
  groupMembershipByWeb: Map<string, Set<number>>,
  pageServerRelativeUrl: string
): boolean {
  const a = mode.access;
  if (a === undefined || a === null) return true;
  const userAllow = new Set(a.allowedUserIds ?? []);
  const groupAllow = new Set(a.allowedGroupIds ?? []);
  if (userAllow.size === 0 && groupAllow.size === 0) return false;
  if (userAllow.has(currentUserId)) return true;
  if (groupAllow.size === 0) return false;
  const key = resolveAccessWebKey(a.webServerRelativeUrl, pageServerRelativeUrl);
  const ms = groupMembershipByWeb.get(key);
  if (!ms) return false;
  const ga = Array.from(groupAllow);
  for (let i = 0; i < ga.length; i++) {
    if (ms.has(ga[i])) return true;
  }
  return false;
}

export function filterViewModesForCurrentUser(
  modes: IListViewModeConfig[],
  currentUserId: number,
  groupMembershipByWeb: Map<string, Set<number>>,
  pageServerRelativeUrl: string
): IListViewModeConfig[] {
  return modes.filter((m) => userCanUseViewMode(m, currentUserId, groupMembershipByWeb, pageServerRelativeUrl));
}

export function pickFallbackViewModeId(
  desiredId: string | undefined,
  visibleModes: IListViewModeConfig[],
  previousFullModes: IListViewModeConfig[]
): string {
  if (visibleModes.length === 0) return desiredId ?? 'all';
  if (desiredId && visibleModes.some((m) => m.id === desiredId)) return desiredId;
  const prev = previousFullModes.find((m) => m.id === desiredId);
  if (prev && visibleModes.some((m) => m.id === prev.id)) return prev.id;
  return visibleModes[0]?.id ?? 'all';
}
