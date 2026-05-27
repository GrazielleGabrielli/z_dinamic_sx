export const DINAMIC_SX_TOOLBAR_CLASS = {
  chromeRow: 'dinamicSxToolbarChromeRow',
  chromeRowStart: 'dinamicSxToolbarChromeRowStart',
  chromeRowEnd: 'dinamicSxToolbarChromeRowEnd',
  end: 'dinamicSxToolbarEnd',
  layoutToggle: 'dinamicSxToolbarLayoutToggle',
  layoutToggleBtn: 'dinamicSxToolbarLayoutToggleBtn',
  primaryBtn: 'dinamicSxToolbarPrimaryBtn',
  defaultBtn: 'dinamicSxToolbarDefaultBtn',
  ghostBtn: 'dinamicSxToolbarGhostBtn',
} as const;

export const DEFAULT_LIST_TOOLBAR_CSS = `
.${DINAMIC_SX_TOOLBAR_CLASS.chromeRow} {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: space-between;
  gap: 12px 8px;
  width: 100%;
  min-width: 0;
}

.${DINAMIC_SX_TOOLBAR_CLASS.chromeRowStart} {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 8px;
  min-width: 0;
  flex: 1 1 auto;
}

.${DINAMIC_SX_TOOLBAR_CLASS.chromeRowEnd} {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: flex-end;
  gap: 8px;
  flex: 0 1 auto;
  margin-left: auto;
}

.${DINAMIC_SX_TOOLBAR_CLASS.end} {
  display: contents;
}

.${DINAMIC_SX_TOOLBAR_CLASS.layoutToggle} {
  display: inline-flex;
  align-items: center;
  background: #f5f5f5;
  border-radius: 8px;
  padding: 5px;
  gap: 2px;
  box-shadow: inset 0 0 0 1px rgba(0, 0, 0, 0.03);
  flex-shrink: 0;
  min-height: 38px;
  box-sizing: border-box;
}

.${DINAMIC_SX_TOOLBAR_CLASS.layoutToggleBtn} {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: 44px;
  height: 28px;
  padding: 0 14px;
  border: none;
  border-radius: 8px;
  cursor: pointer;
  background: transparent;
  color: #525252;
  outline: none;
  box-sizing: border-box;
  font-family: inherit;
  line-height: 1;
  transition: background 0.18s ease, box-shadow 0.18s ease, color 0.18s ease;
}

.${DINAMIC_SX_TOOLBAR_CLASS.layoutToggleBtn} svg {
  display: block;
  width: 18px;
  height: 18px;
  flex-shrink: 0;
}

.${DINAMIC_SX_TOOLBAR_CLASS.layoutToggleBtn}:hover {
  background: rgba(255, 255, 255, 0.55);
  color: #323130;
}

.${DINAMIC_SX_TOOLBAR_CLASS.layoutToggleBtn}[aria-pressed="true"] {
  background: #ffffff;
  color: #242424;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05), 0 2px 8px rgba(0, 0, 0, 0.06);
}

.${DINAMIC_SX_TOOLBAR_CLASS.layoutToggleBtn}[aria-pressed="true"]:hover {
  background: #ffffff;
}

.${DINAMIC_SX_TOOLBAR_CLASS.primaryBtn}.ms-Button {
  height: 38px;
  min-width: 0;
  padding: 0 18px;
  border-radius: 8px;
  border: none;
  font-size: 14px;
  font-weight: 600;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 6px rgba(0, 120, 212, 0.2);
  transition: box-shadow 0.18s ease, transform 0.12s ease;
  flex-shrink: 0;
}

.${DINAMIC_SX_TOOLBAR_CLASS.layoutToggle},
.${DINAMIC_SX_TOOLBAR_CLASS.defaultBtn}.ms-Button,
.${DINAMIC_SX_TOOLBAR_CLASS.ghostBtn}.ms-Button {
  flex-shrink: 0;
}

.${DINAMIC_SX_TOOLBAR_CLASS.primaryBtn}.ms-Button:hover {
  box-shadow: 0 2px 8px rgba(0, 120, 212, 0.28);
}

.${DINAMIC_SX_TOOLBAR_CLASS.primaryBtn}.ms-Button:active {
  transform: translateY(1px);
}

.${DINAMIC_SX_TOOLBAR_CLASS.defaultBtn}.ms-Button {
  height: 38px;
  min-width: 0;
  padding: 0 16px;
  border-radius: 8px;
  border: none;
  background: #ffffff;
  color: #323130;
  font-size: 14px;
  font-weight: 600;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04);
  transition: box-shadow 0.18s ease, background 0.18s ease;
}

.${DINAMIC_SX_TOOLBAR_CLASS.defaultBtn}.ms-Button:hover {
  background: #ffffff;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.07);
}

.${DINAMIC_SX_TOOLBAR_CLASS.ghostBtn}.ms-Button {
  height: 38px;
  padding: 0 14px;
  border-radius: 8px;
  border: none;
  background: #ffffff;
  color: #115ea3;
  font-size: 13px;
  font-weight: 600;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04);
}

.${DINAMIC_SX_TOOLBAR_CLASS.ghostBtn}.ms-Button:hover {
  background: #ffffff;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.07);
  color: #0f6cbd;
}
`.trim();

export function resolveListToolbarCss(customCss?: string): string {
  const custom = (customCss ?? '').trim();
  return custom.length > 0 ? custom : DEFAULT_LIST_TOOLBAR_CSS;
}
