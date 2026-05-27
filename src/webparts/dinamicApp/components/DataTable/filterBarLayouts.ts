export const DINAMIC_SX_FILTER_CLASS = {
  bar: 'dinamicSxFilterBar',
  barFieldsOnly: 'dinamicSxFilterBarFieldsOnly',
  barAdvancedOnly: 'dinamicSxFilterBarAdvancedOnly',
  primary: 'dinamicSxFilterPrimary',
  actions: 'dinamicSxFilterActions',
  headerClear: 'dinamicSxFilterHeaderClear',
  fieldsRow: 'dinamicSxFilterFieldsRow',
  advancedPanel: 'dinamicSxFilterAdvancedPanel',
  control: 'dinamicSxFilterControl',
  label: 'dinamicSxFilterLabel',
  input: 'dinamicSxFilterInput',
  advancedBtn: 'dinamicSxFilterAdvancedBtn',
} as const;

export const DEFAULT_FILTER_BAR_CSS = `
.dinamicSxFilterTableBlock {
  gap: 15px;
}

.dinamicSxToolbarChromeRow + .dinamicSxFilterBar {
  margin-top: -7px;
}

.${DINAMIC_SX_FILTER_CLASS.bar} {
  padding: 12px 14px 14px;
  background: #f5f5f5;
  border: 1px solid #e8e8e8;
  border-radius: 15px;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.03);
}

.${DINAMIC_SX_FILTER_CLASS.barFieldsOnly} {
  padding: 12px 14px 14px;
}

.${DINAMIC_SX_FILTER_CLASS.barAdvancedOnly} {
  padding: 12px 14px 14px;
  box-shadow:
    0 1px 2px rgba(0, 0, 0, 0.03),
    0 3px 10px rgba(0, 0, 0, 0.05);
}

.${DINAMIC_SX_FILTER_CLASS.bar} .${DINAMIC_SX_FILTER_CLASS.advancedPanel} {
  margin-top: 0;
  padding: 0;
  background: transparent;
  border: none;
  border-radius: 0;
  box-shadow: none;
}

.${DINAMIC_SX_FILTER_CLASS.bar} .${DINAMIC_SX_FILTER_CLASS.fieldsRow} + .${DINAMIC_SX_FILTER_CLASS.advancedPanel} {
  margin-top: 10px;
}

.${DINAMIC_SX_FILTER_CLASS.primary} {
  display: flex;
  flex-direction: column;
  gap: 12px;
}

@media (min-width: 900px) {
  .${DINAMIC_SX_FILTER_CLASS.primary} {
    flex-direction: row;
    align-items: flex-end;
    gap: 12px;
  }

  .${DINAMIC_SX_FILTER_CLASS.primary} .${DINAMIC_SX_FILTER_CLASS.fieldsRow} {
    flex: 1 1 auto;
  }
}

.${DINAMIC_SX_FILTER_CLASS.fieldsRow} {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 12px 14px;
  width: 100%;
  min-width: 0;
  align-items: end;
}

@media (max-width: 1279px) {
  .${DINAMIC_SX_FILTER_CLASS.fieldsRow} {
    grid-template-columns: repeat(3, minmax(0, 1fr));
  }
}

@media (max-width: 1023px) {
  .${DINAMIC_SX_FILTER_CLASS.fieldsRow} {
    grid-template-columns: repeat(2, minmax(0, 1fr));
  }
}

@media (max-width: 639px) {
  .${DINAMIC_SX_FILTER_CLASS.fieldsRow} {
    grid-template-columns: minmax(0, 1fr);
    gap: 10px;
  }
}

.${DINAMIC_SX_FILTER_CLASS.actions} {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: flex-end;
  gap: 6px;
  width: 100%;
  flex-shrink: 0;
}

@media (min-width: 900px) {
  .${DINAMIC_SX_FILTER_CLASS.actions} {
    width: auto;
  }
}

.${DINAMIC_SX_FILTER_CLASS.advancedPanel} {
  margin-top: 10px;
  padding: 12px 14px 14px;
  background: #f5f5f5;
  border: 1px solid #e8e8e8;
  border-radius: 15px;
  box-shadow:
    0 1px 2px rgba(0, 0, 0, 0.03),
    0 3px 10px rgba(0, 0, 0, 0.05);
}

.${DINAMIC_SX_FILTER_CLASS.control} {
  display: flex;
  flex-direction: column;
  justify-content: flex-start;
  min-width: 0;
  width: 100%;
}

.${DINAMIC_SX_FILTER_CLASS.label},
.${DINAMIC_SX_FILTER_CLASS.control} label {
  font-size: 11px;
  font-weight: 600;
  color: #605e5c;
  margin: 0 0 5px 2px;
  line-height: 1.3;
  letter-spacing: 0.02em;
}

.${DINAMIC_SX_FILTER_CLASS.input},
.${DINAMIC_SX_FILTER_CLASS.control} input[type="date"],
.${DINAMIC_SX_FILTER_CLASS.control} input[type="text"] {
  width: 100%;
  box-sizing: border-box;
  min-height: 38px;
  padding: 0 13px;
  font-size: 14px;
  font-family: inherit;
  color: #242424;
  background: #ffffff;
  border: 1px solid #e0e0e0;
  border-radius: 15px;
  outline: none;
  box-shadow: none;
  transition: box-shadow 0.18s ease, background 0.18s ease, border-color 0.18s ease;
}

.${DINAMIC_SX_FILTER_CLASS.input}::placeholder,
.${DINAMIC_SX_FILTER_CLASS.control} input::placeholder {
  color: #8a8886;
}

.${DINAMIC_SX_FILTER_CLASS.input}:hover,
.${DINAMIC_SX_FILTER_CLASS.control} input[type="date"]:hover {
  border-color: #d4d4d4;
}

.${DINAMIC_SX_FILTER_CLASS.input}:focus,
.${DINAMIC_SX_FILTER_CLASS.control} input[type="date"]:focus {
  border-color: #c8c8c8;
  background: #ffffff;
  box-shadow: 0 0 0 2px rgba(0, 0, 0, 0.04);
}

.${DINAMIC_SX_FILTER_CLASS.control} .ms-TextField {
  margin: 0;
}

.${DINAMIC_SX_FILTER_CLASS.control} .ms-TextField-fieldGroup {
  min-height: 38px;
  border: 1px solid #e0e0e0;
  border-radius: 15px;
  background: #ffffff;
  box-shadow: none;
  transition: box-shadow 0.18s ease, background 0.18s ease, border-color 0.18s ease;
}

.${DINAMIC_SX_FILTER_CLASS.control} .ms-TextField-fieldGroup:hover {
  border-color: #d4d4d4;
}

.${DINAMIC_SX_FILTER_CLASS.control} .ms-TextField-fieldGroup::after {
  border-radius: 15px;
  border-color: transparent;
}

.${DINAMIC_SX_FILTER_CLASS.control} .ms-TextField-fieldGroup:focus-within {
  border-color: #c8c8c8;
  background: #ffffff;
  box-shadow: 0 0 0 2px rgba(0, 0, 0, 0.04);
}

.${DINAMIC_SX_FILTER_CLASS.control} .ms-TextField-field {
  font-size: 14px;
  color: #242424;
  padding: 0 13px;
}

.${DINAMIC_SX_FILTER_CLASS.control} .ms-TextField-field::placeholder {
  color: #8a8886;
}

.${DINAMIC_SX_FILTER_CLASS.control} .ms-Dropdown {
  margin: 0;
}

.${DINAMIC_SX_FILTER_CLASS.control} .ms-Dropdown-title {
  min-height: 38px;
  line-height: 36px;
  border: 1px solid #e0e0e0;
  border-radius: 15px;
  background: #ffffff;
  font-size: 14px;
  color: #242424;
  padding: 0 34px 0 13px;
  box-shadow: none;
  transition: box-shadow 0.18s ease, background 0.18s ease, border-color 0.18s ease;
}

.${DINAMIC_SX_FILTER_CLASS.control} .ms-Dropdown:hover .ms-Dropdown-title {
  border-color: #d4d4d4;
}

.${DINAMIC_SX_FILTER_CLASS.control} .ms-Dropdown:focus::after {
  border-radius: 15px;
  border-color: transparent;
}

.${DINAMIC_SX_FILTER_CLASS.control} .ms-Dropdown.is-open .ms-Dropdown-title,
.${DINAMIC_SX_FILTER_CLASS.control} .ms-Dropdown:focus .ms-Dropdown-title {
  border-color: #c8c8c8;
  box-shadow: 0 0 0 2px rgba(0, 0, 0, 0.04);
}

.${DINAMIC_SX_FILTER_CLASS.control} .ms-Dropdown-caretDownWrapper {
  color: #605e5c;
  right: 11px;
}

.${DINAMIC_SX_FILTER_CLASS.advancedBtn} {
  height: 38px !important;
  padding: 0 14px !important;
  border-radius: 15px !important;
  border: none !important;
  background: #ffffff !important;
  color: #323130 !important;
  font-size: 13px !important;
  font-weight: 600 !important;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04) !important;
  transition: box-shadow 0.18s ease, background 0.18s ease !important;
  flex-shrink: 0;
}

.${DINAMIC_SX_FILTER_CLASS.advancedBtn}:hover {
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.07) !important;
}

.${DINAMIC_SX_FILTER_CLASS.advancedBtn} i {
  color: #605e5c;
}

.${DINAMIC_SX_FILTER_CLASS.advancedBtn}.dinamicSxFilterAdvancedBtn--open {
  background: #ffffff !important;
  color: #242424 !important;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05), 0 2px 10px rgba(0, 0, 0, 0.08) !important;
}

.${DINAMIC_SX_FILTER_CLASS.advancedBtn}.dinamicSxFilterAdvancedBtn--open i {
  color: #323130;
}

.${DINAMIC_SX_FILTER_CLASS.headerClear} {
  height: 38px !important;
  padding: 0 14px !important;
  border-radius: 15px !important;
  border: none !important;
  background: #ffffff !important;
  color: #a4262c !important;
  font-size: 13px !important;
  font-weight: 600 !important;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04) !important;
  transition: box-shadow 0.18s ease !important;
  flex-shrink: 0;
}

.${DINAMIC_SX_FILTER_CLASS.headerClear}:hover {
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.07) !important;
}

.${DINAMIC_SX_FILTER_CLASS.headerClear} i {
  color: #a4262c;
}
`.trim();

export function resolveFilterBarCss(customCss: string | undefined): string {
  const custom = (customCss ?? '').trim();
  return custom.length > 0 ? custom : DEFAULT_FILTER_BAR_CSS;
}
