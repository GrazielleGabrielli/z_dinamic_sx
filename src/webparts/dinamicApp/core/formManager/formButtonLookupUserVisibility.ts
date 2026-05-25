import type { IFormCustomButtonConfig } from '../config/types/formManager';

function tryGetObjectProp(obj: Record<string, unknown>, key: string): unknown {
  if (key in obj) return obj[key];
  return undefined;
}

function readPathValue(root: unknown, path: string[]): unknown {
  if (!path.length) return root;
  if (root === null || root === undefined) return undefined;
  if (Array.isArray(root)) {
    const out: unknown[] = [];
    for (let i = 0; i < root.length; i++) {
      const v = readPathValue(root[i], path);
      if (v !== undefined && v !== null) {
        if (Array.isArray(v)) out.push(...v);
        else out.push(v);
      }
    }
    return out.length ? out : undefined;
  }
  if (typeof root === 'object') {
    const obj = root as Record<string, unknown>;
    const head = path[0];
    const next = tryGetObjectProp(obj, head);
    return readPathValue(next, path.slice(1));
  }
  return undefined;
}

export function normalizeLookupUserFieldPath(raw: string): string {
  return raw
    .split('/')
    .map((s) => s.trim())
    .filter(Boolean)
    .join('/');
}

export function userFieldIdsFromValue(v: unknown): number[] {
  if (v === null || v === undefined) return [];
  if (typeof v === 'number' && isFinite(v)) return v > 0 ? [v] : [];
  if (Array.isArray(v)) {
    const out: number[] = [];
    for (let i = 0; i < v.length; i++) {
      out.push(...userFieldIdsFromValue(v[i]));
    }
    return out.filter((x, j, a) => a.indexOf(x) === j);
  }
  if (typeof v === 'object' && v !== null && 'Id' in v) {
    const id = (v as Record<string, unknown>).Id;
    if (typeof id === 'number' && isFinite(id) && id > 0) return [id];
  }
  return [];
}

export function readLookupUserFieldValue(
  fieldPath: string,
  values: Record<string, unknown>,
  lookupOptionSnapshots?: Readonly<
    Record<string, Record<string, unknown> | Record<string, unknown>[] | undefined>
  >
): unknown {
  const trimmed = normalizeLookupUserFieldPath(fieldPath);
  if (!trimmed) return undefined;
  const segments = trimmed.split('/');
  const rootName = segments[0];
  if (segments.length === 1) return values[rootName];
  const subPath = segments.slice(1);
  const fromRoot = readPathValue(values[rootName], subPath);
  if (fromRoot !== undefined && fromRoot !== null) return fromRoot;
  const snap = lookupOptionSnapshots?.[rootName];
  if (snap === undefined || snap === null) return fromRoot;
  return readPathValue(snap, subPath);
}

export function currentUserMatchesLookupUserFieldPath(
  fieldPath: string,
  currentUserId: number,
  values: Record<string, unknown>,
  lookupOptionSnapshots?: Readonly<
    Record<string, Record<string, unknown> | Record<string, unknown>[] | undefined>
  >
): boolean {
  if (!currentUserId || currentUserId <= 0) return false;
  const val = readLookupUserFieldValue(fieldPath, values, lookupOptionSnapshots);
  const ids = userFieldIdsFromValue(val);
  return ids.indexOf(currentUserId) !== -1;
}

export function userInAnyLookupUserField(
  currentUserId: number,
  values: Record<string, unknown>,
  paths: string[] | undefined,
  lookupOptionSnapshots?: Readonly<
    Record<string, Record<string, unknown> | Record<string, unknown>[] | undefined>
  >
): boolean {
  if (!paths || paths.length === 0) return true;
  for (let i = 0; i < paths.length; i++) {
    const p = normalizeLookupUserFieldPath(paths[i]);
    if (!p) continue;
    if (currentUserMatchesLookupUserFieldPath(p, currentUserId, values, lookupOptionSnapshots)) return true;
  }
  return false;
}

export function ruleAppliesLookupUserFieldFilters(
  currentUserId: number,
  values: Record<string, unknown>,
  rule: Pick<IFormCustomButtonConfig, 'lookupUserFieldPaths' | 'excludeLookupUserFieldPaths'>,
  lookupOptionSnapshots?: Readonly<
    Record<string, Record<string, unknown> | Record<string, unknown>[] | undefined>
  >
): boolean {
  if (!userInAnyLookupUserField(currentUserId, values, rule.lookupUserFieldPaths, lookupOptionSnapshots)) {
    return false;
  }
  const ex = rule.excludeLookupUserFieldPaths;
  if (ex && ex.length && userInAnyLookupUserField(currentUserId, values, ex, lookupOptionSnapshots)) {
    return false;
  }
  return true;
}

export function collectLookupSubfieldsFromUserVisibilityPaths(
  paths: readonly string[] | undefined,
  lookupInternalName: string,
  into: Set<string>
): void {
  const root = lookupInternalName.trim();
  if (!root || !paths?.length) return;
  for (let i = 0; i < paths.length; i++) {
    const parts = normalizeLookupUserFieldPath(paths[i]).split('/');
    if (parts.length >= 2 && parts[0] === root && parts[1]) into.add(parts[1]);
  }
}
