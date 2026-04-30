import type { IGroupDetails } from './types';

export function isExcludedNativeSharePointSiteGroupTitle(title: string | undefined): boolean {
  const t = (title ?? '').trim();
  if (!t) return true;
  if (/^SharingLinks\./i.test(t)) return true;
  if (/Limited Access System/i.test(t)) return true;
  return false;
}

export function filterSiteGroupsForPicker(groups: IGroupDetails[]): IGroupDetails[] {
  return groups.filter((g) => !isExcludedNativeSharePointSiteGroupTitle(g.Title));
}

export function filterSiteGroupsByNameQuery(groups: IGroupDetails[], query: string): IGroupDetails[] {
  const q = query.trim().toLowerCase();
  if (!q) return groups;
  return groups.filter((g) => (g.Title ?? '').toLowerCase().includes(q));
}
