import type { ITableLayoutCssSlots } from './types';

/** Slots de exemplo para testar estilos na aba Layout; remova ou esvazie em produção. */
export const DEFAULT_TABLE_LAYOUT_TEST_CSS_SLOTS: ITableLayoutCssSlots = {
  viewRoot: `background: linear-gradient(135deg, #e0e7ff 0%, #fae8ff 100%);
padding: 12px;
border-radius: 10px;
border: 2px solid #6366f1;`,
  toolbar: `background: #312e81;
padding: 10px 12px;
border-radius: 8px;
gap: 12px;`,
  scrollWrap: `background: #f1f5f9;
border: 2px dashed #64748b;
border-radius: 6px;
max-height: 80px;`,
  table: `outline: 3px solid #0ea5e9;
background: #ecfeff;`,
  thead: `background: #1e293b;
color: #f8fafc;`,
  headerRow: `border-bottom: 4px solid #f59e0b;`,
  headerCell: `background: #059669 !important;
color: #fff !important;
padding: 12px !important;
font-weight: 700 !important;`,
  headerCellInner: `background: rgba(99, 102, 241, 0.25);
padding: 4px 8px;
border-radius: 4px;`,
  headerFilterTrigger: `background: #fef08a;
border-radius: 4px;
padding: 2px 6px !important;
transform: scale(1.2);`,
  body: `background: #fff7ed;
outline: 2px solid #ea580c;`,
  row: `background: #dbeafe !important;
border-left: 5px solid #2563eb !important;`,
  cell: `color: #7c3aed !important;
font-weight: 600 !important;
border: 1px dotted #a78bfa !important;`,
  empty: `background: #fce7f3 !important;
border-radius: 12px !important;
outline: 2px solid #db2777;`,
  loading: `background: #e0f2fe !important;
border: 2px solid #0284c7;
border-radius: 8px;`,
  error: `outline: 3px solid #dc2626;
padding: 4px !important;
background: #fef2f2 !important;`,
  pagination: `background: #ecfdf5;
padding: 10px;
border-radius: 8px;
border: 2px solid #10b981;`,
};

export const DEFAULT_TABLE_LAYOUT_TEST_CSS_FREE = `.dinamicSxTableRow:nth-child(even) .dinamicSxTableCell {
  background: #f8fafc !important;
}
.dinamicSxTableHeaderCell[data-field="Title"] {
  background: #422006 !important;
  color: #ffedd5 !important;
}
`;
