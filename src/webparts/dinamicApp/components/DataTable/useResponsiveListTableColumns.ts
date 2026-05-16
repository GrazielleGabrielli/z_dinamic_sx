import { useEffect, useMemo, useState } from 'react';
import type { ITableColumnConfig } from '../../core/table/types';
import {
  listViewBreakpointIndexForWidth,
  resolveListColumnSpanAtBreakpoint,
} from '../../core/listView/listViewColumnBreakpoints';

function withResolvedWidths(columns: ITableColumnConfig[], breakpointIndex: number): ITableColumnConfig[] {
  return columns.map((c) => {
    const map = c.columnSpanByBreakpoint;
    if (!map || Object.keys(map).length === 0) return c;
    const span = resolveListColumnSpanAtBreakpoint(breakpointIndex, map);
    const pct = Math.round((span / 12) * 10000) / 100;
    return { ...c, width: `${pct}%` };
  });
}

export function useResponsiveListTableColumns(columns: ITableColumnConfig[]): ITableColumnConfig[] {
  const [breakpointIndex, setBreakpointIndex] = useState(() =>
    typeof window !== 'undefined' ? listViewBreakpointIndexForWidth(window.innerWidth) : 0
  );
  useEffect(() => {
    const onResize = (): void => setBreakpointIndex(listViewBreakpointIndexForWidth(window.innerWidth));
    onResize();
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, []);
  return useMemo(() => withResolvedWidths(columns, breakpointIndex), [columns, breakpointIndex]);
}
