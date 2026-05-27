import type { IListViewModeConfig, TViewModePicker } from '../../core/config/types';

export const VIEW_MODE_PICKER_OPTIONS: { key: TViewModePicker; text: string }[] = [
  { key: 'dropdown', text: 'Lista suspensa (select)' },
  { key: 'buttons', text: 'Botões' },
  { key: 'tabs', text: 'Abas com sublinhado' },
  { key: 'segmented', text: 'Segmentado (pílulas no trilho)' },
  { key: 'pills', text: 'Pílulas com contador' },
  { key: 'iconsUnderline', text: 'Ícones e sublinhado' },
  { key: 'iconsBadges', text: 'Ícones e contador' },
];

const PICKER_KEYS = new Set<TViewModePicker>(VIEW_MODE_PICKER_OPTIONS.map((o) => o.key));

export function normalizeViewModePicker(raw: unknown): TViewModePicker {
  if (typeof raw === 'string' && PICKER_KEYS.has(raw as TViewModePicker)) {
    return raw as TViewModePicker;
  }
  return 'dropdown';
}

export function defaultViewModeIcon(mode: IListViewModeConfig): string {
  if (mode.iconName?.trim()) return mode.iconName.trim();
  if (mode.id === 'all') return 'ViewList';
  if (mode.id === 'mine') return 'Contact';
  return 'Filter';
}

export const DEFAULT_VIEW_MODE_CSS = `
.dinamicSxViewModeBar {
  flex: 0 1 auto;
  min-width: 0;
}

.dinamicSxViewModeDropdown .ms-Dropdown {
  max-width: 240px;
}

.dinamicSxViewModeButtonsRow {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  align-items: center;
}

.dinamicSxViewModeTab.ms-Button {
  border-radius: 8px;
}

.dinamicSxViewModeDropdown .ms-Dropdown-title {
  border-radius: 8px;
}

.dinamicSxViewModeSegmentedTrack {
  display: inline-flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 2px;
  padding: 5px;
  background: #f5f5f5;
  border-radius: 8px;
  max-width: 100%;
  border: none;
  box-shadow: inset 0 0 0 1px rgba(0, 0, 0, 0.03);
}

.dinamicSxViewModeSegmentedItem {
  display: inline-flex;
  align-items: center;
  gap: 7px;
  padding: 9px 16px;
  border: none;
  border-radius: 8px;
  background: transparent;
  color: #525252;
  font-size: 14px;
  font-weight: 500;
  font-family: inherit;
  cursor: pointer;
  transition: background 0.18s ease, box-shadow 0.18s ease, color 0.18s ease;
  white-space: nowrap;
}

.dinamicSxViewModeSegmentedItem:hover {
  color: #323130;
  background: rgba(255, 255, 255, 0.55);
}

.dinamicSxViewModeSegmentedItem[aria-selected="true"] {
  background: #ffffff;
  color: #242424;
  font-weight: 600;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05), 0 2px 8px rgba(0, 0, 0, 0.06);
}

.dinamicSxViewModeSegmentedItem[aria-selected="true"]:hover {
  background: #ffffff;
}

.dinamicSxViewModeSegmentedItem i {
  color: #757575;
  font-size: 15px;
}

.dinamicSxViewModeSegmentedItem[aria-selected="true"] i {
  color: #424242;
}

.dinamicSxViewModeUnderlineTrack {
  display: flex;
  flex-wrap: wrap;
  align-items: stretch;
  gap: 0;
  max-width: 100%;
  border-bottom: 1px solid #e4e4e7;
}

.dinamicSxViewModeUnderlineItem {
  flex: 0 0 auto;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 12px 16px;
  border: none;
  background: transparent;
  color: #71717a;
  font-size: 14px;
  font-weight: 500;
  font-family: inherit;
  cursor: pointer;
  position: relative;
  white-space: nowrap;
  transition: color 0.15s ease;
}

.dinamicSxViewModeUnderlineItem:hover {
  color: #3f3f46;
}

.dinamicSxViewModeUnderlineItem[aria-selected="true"] {
  color: #0f6cbd;
  font-weight: 600;
}

.dinamicSxViewModeUnderlineItem[aria-selected="true"]::after {
  content: "";
  position: absolute;
  left: 12px;
  right: 12px;
  bottom: 0;
  height: 2px;
  background: #0f6cbd;
  border-radius: 2px 2px 0 0;
}

.dinamicSxViewModePillsRow {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 20px;
}

.dinamicSxViewModePill {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 4px 0;
  border: none;
  background: transparent;
  font-family: inherit;
  font-size: 14px;
  font-weight: 500;
  color: #3f3f46;
  cursor: pointer;
  transition: color 0.15s ease;
}

.dinamicSxViewModePill[aria-selected="true"] {
  color: #0f6cbd;
}

.dinamicSxViewModePillBadge {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: 22px;
  padding: 2px 8px;
  border-radius: 8px;
  font-size: 12px;
  font-weight: 600;
  line-height: 1.3;
  background: #f4f4f5;
  color: #52525b;
}

.dinamicSxViewModePill[aria-selected="true"] .dinamicSxViewModePillBadge {
  background: #e8f3fc;
  color: #0f6cbd;
}

.dinamicSxViewModePillBadge--success {
  background: #dcfce7;
  color: #166534;
}

.dinamicSxViewModePillBadge--muted {
  background: transparent;
  color: #71717a;
  padding-left: 0;
  min-width: 0;
}

.dinamicSxViewModeIconRow {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 4px 20px;
  border-bottom: 1px solid #e4e4e7;
  padding-bottom: 2px;
}

.dinamicSxViewModeIconItem {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 10px 4px 12px;
  border: none;
  background: transparent;
  font-family: inherit;
  font-size: 14px;
  font-weight: 500;
  color: #71717a;
  cursor: pointer;
  position: relative;
  white-space: nowrap;
  transition: color 0.15s ease;
}

.dinamicSxViewModeIconItem i {
  font-size: 16px;
  color: #a1a1aa;
}

.dinamicSxViewModeIconItem[aria-selected="true"] {
  color: #18181b;
}

.dinamicSxViewModeIconItem[aria-selected="true"] i {
  color: #0f6cbd;
}

.dinamicSxViewModeIconItem[aria-selected="true"]::after {
  content: "";
  position: absolute;
  left: 0;
  right: 0;
  bottom: 0;
  height: 2px;
  background: #18181b;
  border-radius: 2px 2px 0 0;
}

.dinamicSxViewModeIconItem .dinamicSxViewModePillBadge {
  margin-left: 2px;
}
`.trim();

export function resolveViewModeCss(customCss: string | undefined): string {
  const custom = (customCss ?? '').trim();
  return custom.length > 0 ? custom : DEFAULT_VIEW_MODE_CSS;
}

export function pillBadgeVariant(mode: IListViewModeConfig, selected: boolean): string {
  if (mode.badgeCount === undefined) return 'muted';
  if (selected) return '';
  if (mode.id === 'mine' || mode.label.toLowerCase().indexOf('conclu') !== -1) return 'success';
  return '';
}
