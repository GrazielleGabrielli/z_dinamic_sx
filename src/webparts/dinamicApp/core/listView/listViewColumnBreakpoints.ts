import {
  LIST_VIEW_COLUMN_BREAKPOINT_KEYS,
  type TListViewColumnBreakpoint,
} from '../config/types/listViewBreakpoints';

export const LIST_VIEW_COLUMN_BREAKPOINT_ORDER = [
  ...LIST_VIEW_COLUMN_BREAKPOINT_KEYS,
] as readonly TListViewColumnBreakpoint[];

export const LIST_VIEW_COLUMN_BREAKPOINT_MIN_PX: Record<TListViewColumnBreakpoint, number> = {
  xs: 0,
  s: 480,
  m: 640,
  l: 1024,
  xl: 1366,
  xxl: 1920,
};

export const LIST_VIEW_COLUMN_BREAKPOINT_LABEL: Record<TListViewColumnBreakpoint, string> = {
  xs: 'XS',
  s: 'S',
  m: 'M',
  l: 'L',
  xl: 'XL',
  xxl: 'XXL',
};

const ALLOWED_LIST_SPANS = new Set([3, 4, 6, 8, 12]);

export function clampListColumnSpan(n: number): number {
  const x = Math.round(Number(n));
  return ALLOWED_LIST_SPANS.has(x) ? x : 12;
}

export function listViewBreakpointIndexForWidth(width: number): number {
  let idx = 0;
  for (let i = 0; i < LIST_VIEW_COLUMN_BREAKPOINT_ORDER.length; i++) {
    const bp = LIST_VIEW_COLUMN_BREAKPOINT_ORDER[i];
    if (width >= LIST_VIEW_COLUMN_BREAKPOINT_MIN_PX[bp]) idx = i;
  }
  return idx;
}

export function resolveListColumnSpanAtBreakpoint(
  breakpointIndex: number,
  map: Partial<Record<TListViewColumnBreakpoint, number>> | undefined
): number {
  if (!map) return 12;
  const i0 = Math.max(0, Math.min(breakpointIndex, LIST_VIEW_COLUMN_BREAKPOINT_ORDER.length - 1));
  for (let i = i0; i >= 0; i--) {
    const k = LIST_VIEW_COLUMN_BREAKPOINT_ORDER[i];
    const v = map[k];
    if (v !== undefined) return clampListColumnSpan(v);
  }
  return 12;
}

export function listColumnHasResponsiveSpan(
  map: Partial<Record<TListViewColumnBreakpoint, number>> | undefined
): boolean {
  return !!(map && Object.keys(map).length > 0);
}
