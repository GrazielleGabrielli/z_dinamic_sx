export const LIST_VIEW_COLUMN_BREAKPOINT_KEYS = ['xs', 's', 'm', 'l', 'xl', 'xxl'] as const;
export type TListViewColumnBreakpoint = (typeof LIST_VIEW_COLUMN_BREAKPOINT_KEYS)[number];
