import type { ITableLayoutCssSlots, ITableRowStyleRule, TTableCssSlot } from '../../core/config/types';
import { toTableRowRuleDataToken } from '../../core/table/utils/tableRowStyleRuleEval';

export const DINAMIC_SX_TABLE_CLASS = {
  viewRoot: 'dinamicSxTableView',
  toolbar: 'dinamicSxTableToolbar',
  scrollWrap: 'dinamicSxTableScroll',
  table: 'dinamicSxTableElement',
  thead: 'dinamicSxTableHead',
  headerRow: 'dinamicSxTableHeaderRow',
  headerCell: 'dinamicSxTableHeaderCell',
  headerCellInner: 'dinamicSxTableHeaderCellInner',
  headerFilterTrigger: 'dinamicSxTableHeaderFilterTrigger',
  body: 'dinamicSxTableBody',
  row: 'dinamicSxTableRow',
  cell: 'dinamicSxTableCell',
  empty: 'dinamicSxTableEmpty',
  loading: 'dinamicSxTableLoading',
  error: 'dinamicSxTableError',
  pagination: 'dinamicSxTablePagination',
  paginationBtn: 'dinamicSxTablePaginationBtn',
  columnFilter: 'dinamicSxTableColumnFilter',
} as const;

export const DINAMIC_SX_CARD_CLASS = {
  toolbar: 'dinamicSxCardToolbar',
  toolbarInner: 'dinamicSxCardToolbarInner',
  toolbarItem: 'dinamicSxCardToolbarItem',
  toolbarItemLabel: 'dinamicSxCardToolbarItemLabel',
  toolbarFilterTrigger: 'dinamicSxCardToolbarFilterTrigger',
  grid: 'dinamicSxCardGrid',
  card: 'dinamicSxCard',
  cardClickable: 'dinamicSxCard--clickable',
  cardHeader: 'dinamicSxCardHeader',
  cardBody: 'dinamicSxCardBody',
  title: 'dinamicSxCardTitle',
  fieldRow: 'dinamicSxCardField',
  fieldLabel: 'dinamicSxCardLabel',
  fieldValue: 'dinamicSxCardValue',
  actions: 'dinamicSxCardActions',
} as const;

export type TDinamicSxTableClassKey = keyof typeof DINAMIC_SX_TABLE_CLASS;

export const TABLE_LAYOUT_EDITOR_ROWS: readonly {
  slot: TTableCssSlot;
  title: string;
  hint: string;
}[] = [
  {
    slot: 'viewRoot',
    title: 'Container da lista',
    hint:
      'Envolve toolbar, tabela e paginação. Use para margem externa, fundo geral, borda ou sombra de todo o bloco da lista.',
  },
  {
    slot: 'toolbar',
    title: 'Barra superior',
    hint: 'Área do seletor de visualização e do botão Exportar PDF. Ajuste alinhamento, espaçamento entre controles ou fundo dessa faixa.',
  },
  {
    slot: 'scrollWrap',
    title: 'Área de rolagem',
    hint: 'Div ao redor da tabela com scroll horizontal. Útil para borda, raio de canto ou limite de altura quando há muitas colunas.',
  },
  {
    slot: 'table',
    title: 'Tabela (<table>)',
    hint: 'Largura, borda externa da grade, collapse/separate e tipografia base herdada pelas células.',
  },
  {
    slot: 'thead',
    title: 'Bloco do cabeçalho (<thead>)',
    hint: 'Fundo ou borda comum a todo o cabeçalho. Combinado com células <th> para estilo da primeira linha.',
  },
  {
    slot: 'headerRow',
    title: 'Linha do cabeçalho (<tr>)',
    hint: 'Altura mínima, borda inferior da linha de títulos ou efeito quando o cabeçalho é sticky.',
  },
  {
    slot: 'headerCell',
    title: 'Células do cabeçalho (<th>)',
    hint:
      'Padding, fonte, cor e borda de cada coluna. Use o atributo [data-field="NomeInterno"] no seletor para uma coluna específica (ex.: Title).',
  },
  {
    slot: 'headerCellInner',
    title: 'Conteúdo do cabeçalho',
    hint: 'Span que agrupa rótulo, filtro e ordenação. Ajuste gap, alinhamento vertical ou tamanho dos ícones via filhos.',
  },
  {
    slot: 'headerFilterTrigger',
    title: 'Ícone de filtro',
    hint: 'Área clicável do filtro. Pode alterar opacidade, margem ou cursor; o botão interno é do Fluent UI.',
  },
  {
    slot: 'body',
    title: 'Corpo (<tbody>)',
    hint: 'Fundo ou espaçamento global das linhas de dados antes de estilizar cada <tr> individualmente.',
  },
  {
    slot: 'row',
    title: 'Linhas de dados (<tr>)',
    hint: 'Borda entre linhas, cores zebradas (:nth-child), hover ou altura mínima da linha.',
  },
  {
    slot: 'cell',
    title: 'Células de dados (<td>)',
    hint:
      'Padding, cor do texto e borda da célula. Para uma coluna: .dinamicSxTableCell[data-field="Campo"] { ... } no bloco CSS adicional abaixo.',
  },
  {
    slot: 'empty',
    title: 'Lista vazia',
    hint: 'Mensagem quando não há itens ou não há colunas. Padding, fundo e cor do texto do estado vazio.',
  },
  {
    slot: 'loading',
    title: 'Carregando',
    hint: 'Área exibida enquanto os dados carregam. Estilo do spinner/texto (o spinner em si é do Fluent).',
  },
  {
    slot: 'error',
    title: 'Erro',
    hint: 'Container da mensagem de erro ao falhar a consulta. Margem e layout ao redor do MessageBar.',
  },
  {
    slot: 'pagination',
    title: 'Paginação',
    hint: 'Botões Anterior/Próxima, números de página e texto “Página X”. Estilize botões e alinhamento da barra.',
  },
] as const;

export const TABLE_LAYOUT_EDITOR_GROUPS: readonly {
  id: string;
  label: string;
  blurb: string;
  slots: readonly TTableCssSlot[];
}[] = [
  {
    id: 'shell',
    label: 'Estrutura da lista',
    blurb: 'Contêiner externo, barra de ferramentas, rolagem e o elemento table.',
    slots: ['viewRoot', 'toolbar', 'scrollWrap', 'table'],
  },
  {
    id: 'header',
    label: 'Cabeçalho da tabela',
    blurb: 'Bloco thead, linha e células de título, área do rótulo e do filtro.',
    slots: ['thead', 'headerRow', 'headerCell', 'headerCellInner', 'headerFilterTrigger'],
  },
  {
    id: 'grid',
    label: 'Corpo da grade',
    blurb: 'Tbody, linhas de dados e células.',
    slots: ['body', 'row', 'cell'],
  },
  {
    id: 'states',
    label: 'Estados da lista',
    blurb: 'Quando não há dados, durante o carregamento ou em caso de erro.',
    slots: ['empty', 'loading', 'error'],
  },
  {
    id: 'pagination',
    label: 'Paginação',
    blurb: 'Controles abaixo da tabela.',
    slots: ['pagination'],
  },
] as const;

export const TABLE_LAYOUT_SLOT_ORDER: TTableCssSlot[] = TABLE_LAYOUT_EDITOR_ROWS.map((r) => r.slot);

export function sanitizeTableCssSlots(raw: unknown): ITableLayoutCssSlots | undefined {
  if (raw === undefined || raw === null) return undefined;
  if (typeof raw !== 'object' || Array.isArray(raw)) return undefined;
  const src = raw as Record<string, unknown>;
  const out: ITableLayoutCssSlots = {};
  for (let i = 0; i < TABLE_LAYOUT_SLOT_ORDER.length; i++) {
    const slot = TABLE_LAYOUT_SLOT_ORDER[i];
    const v = src[slot];
    if (typeof v === 'string' && v.trim().length > 0) {
      out[slot] = v;
    }
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

export function mergeCustomTableCss(
  slots: ITableLayoutCssSlots | undefined,
  legacyFreeform: string | undefined
): string {
  const parts: string[] = [];
  if (slots) {
    for (let i = 0; i < TABLE_LAYOUT_SLOT_ORDER.length; i++) {
      const slot = TABLE_LAYOUT_SLOT_ORDER[i];
      const body = slots[slot]?.trim();
      if (!body) continue;
      const cls = DINAMIC_SX_TABLE_CLASS[slot];
      parts.push(`.${cls} {\n${body}\n}`);
    }
  }
  const free = (legacyFreeform ?? '').trim();
  if (free) parts.push(free);
  return parts.join('\n\n').trim();
}

function buildDefaultTableLayoutCss(): string {
  const c = DINAMIC_SX_TABLE_CLASS;
  return `
.${c.viewRoot} {
  margin-top: 2px;
  font-family: "Segoe UI Variable", "Segoe UI", -apple-system, BlinkMacSystemFont, "Helvetica Neue", Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
}

.${c.toolbar} {
  margin-bottom: 8px;
  padding: 0;
}

.${c.scrollWrap} {
  overflow-x: auto;
  overflow-y: hidden;
  background: #ffffff;
  border: 1px solid #e8e8e8;
  border-radius: 8px;
  box-shadow:
    0 1px 2px rgba(0, 0, 0, 0.03),
    0 3px 10px rgba(0, 0, 0, 0.05);
}

.${c.table} {
  width: 100%;
  border-collapse: separate;
  border-spacing: 0;
  font-size: 14px;
  line-height: 1.5;
  color: #242424;
  font-variant-numeric: tabular-nums;
}

.${c.thead} {
  position: sticky;
  top: 0;
  z-index: 4;
}

.${c.headerRow} {
  background: #f5f5f5;
  box-shadow: inset 0 -1px 0 #e8e8e8;
}

.${c.headerCell} {
  padding: 0 16px;
  height: 44px;
  font-size: 12px;
  font-weight: 600;
  text-transform: none;
  letter-spacing: 0.02em;
  color: #605e5c;
  white-space: nowrap;
  vertical-align: middle;
  border: none;
  background: transparent;
}

.${c.headerCellInner} {
  display: inline-flex;
  align-items: center;
  gap: 6px;
}

.${c.headerFilterTrigger} {
  opacity: 0.6;
  transition: opacity 0.18s ease;
}

.${c.headerFilterTrigger}:hover {
  opacity: 1;
}

.${c.body} {
  background: #ffffff;
}

.${c.row} {
  background: #ffffff;
  transition: background-color 0.16s ease;
}

.${c.row}:hover {
  background: #fafafa;
}

.${c.cell} {
  padding: 12px 16px;
  border-bottom: 1px solid #ececec;
  vertical-align: middle;
  color: #323130;
}

.${c.cell}[data-field="Title"],
.${c.cell}[data-field="LinkTitle"] {
  color: #242424;
  font-weight: 600;
}

.${c.row}:last-child .${c.cell} {
  border-bottom: none;
}

.${c.empty},
.${c.loading} {
  padding: 40px 24px;
  text-align: center;
  color: #605e5c;
  font-size: 14px;
  background: #f5f5f5;
  border-radius: 8px;
  border: 1px solid #e8e8e8;
  margin-top: 8px;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.03);
}

.${c.error} {
  margin-top: 10px;
}

.${c.pagination} {
  margin-top: 12px;
  padding: 2px 0;
  gap: 8px;
}

.${c.paginationBtn} {
  font-family: inherit;
  font-size: 14px;
  font-weight: 600;
  line-height: 1.2;
  color: #323130;
  background: #ffffff;
  border: 1px solid #e0e0e0;
  border-radius: 8px;
  padding: 0 16px;
  min-height: 38px;
  cursor: pointer;
  transition: background-color 0.15s ease, border-color 0.15s ease, box-shadow 0.15s ease, color 0.15s ease;
  box-shadow: none;
}

.${c.paginationBtn}:hover {
  background: #ffffff;
  border-color: #d4d4d4;
  color: #242424;
}

.${c.paginationBtn}:active {
  background: #f5f5f5;
}

.${c.paginationBtn}[data-active="true"] {
  border-color: #0f6cbd;
  color: #0f6cbd;
  background: #f0f6fc;
}

.${c.pagination} > span {
  font-size: 13px;
  color: #605e5c;
  font-weight: 500;
}

.dinamicSxTablePagination--compact .${c.paginationBtn} {
  padding: 0 12px;
  min-height: 34px;
  font-size: 13px;
  border-radius: 8px;
}

.${c.columnFilter} .ms-Button--primary {
  border: none;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 6px rgba(0, 120, 212, 0.2);
}
`.trim();
}

/** Callout de filtro por coluna renderiza no portal do Fluent (fora do escopo da instância). */
export const TABLE_COLUMN_FILTER_PORTAL_CSS = (() => {
  const c = DINAMIC_SX_TABLE_CLASS;
  return `
.${c.columnFilter} {
  padding: 12px 14px;
  min-width: 240px;
  box-sizing: border-box;
}

.${c.columnFilter} .ms-TextField-fieldGroup {
  min-height: 38px;
  border: 1px solid #e0e0e0;
  border-radius: 8px;
  background: #ffffff;
  box-shadow: none;
}

.${c.columnFilter} .ms-TextField-fieldGroup:hover {
  border-color: #d4d4d4;
}

.${c.columnFilter} .ms-TextField-fieldGroup:focus-within {
  border-color: #c8c8c8;
  box-shadow: 0 0 0 2px rgba(0, 0, 0, 0.04);
}

.${c.columnFilter} .ms-TextField-field {
  font-size: 14px;
  padding: 0 13px;
}

.${c.columnFilter} .ms-Button {
  border-radius: 8px;
  min-height: 38px;
  font-weight: 600;
}

.${c.columnFilter} .ms-Button--default {
  border: 1px solid #e0e0e0;
  background: #ffffff;
}

.${c.columnFilter} .ms-Button--primary {
  border: none;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 6px rgba(0, 120, 212, 0.2);
}
`.trim();
})();

export const DEFAULT_TABLE_LAYOUT_CSS = buildDefaultTableLayoutCss();

function buildDefaultCardLayoutCss(): string {
  const c = DINAMIC_SX_CARD_CLASS;
  const t = DINAMIC_SX_TABLE_CLASS;
  return `
.${c.toolbar} {
  padding: 10px 14px;
  background: #f5f5f5;
  border-bottom: 1px solid #e8e8e8;
  overflow-x: auto;
}

.${c.toolbarInner} {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  align-items: center;
  min-width: min-content;
}

.${c.toolbarItem} {
  display: inline-flex;
  align-items: center;
  gap: 2px;
  padding: 2px 4px 2px 10px;
  background: #ffffff;
  border: 1px solid #e0e0e0;
  border-radius: 8px;
  min-height: 32px;
  box-sizing: border-box;
}

.${c.toolbarItemLabel} {
  font-size: 12px;
  font-weight: 600;
  color: #605e5c;
  white-space: nowrap;
  line-height: 1.2;
}

.${c.toolbarFilterTrigger} {
  opacity: 0.65;
  transition: opacity 0.16s ease;
}

.${c.toolbarFilterTrigger}:hover {
  opacity: 1;
}

.${c.grid} {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
  gap: 16px;
  padding: 16px;
  box-sizing: border-box;
}

.${c.card} {
  display: flex;
  flex-direction: column;
  min-width: 0;
  background: #ffffff;
  border: 1px solid #ececec;
  border-radius: 10px;
  overflow: hidden;
  transition:
    border-color 0.16s ease,
    box-shadow 0.16s ease,
    transform 0.16s ease;
}

.${c.card}:hover {
  border-color: #d4d4d4;
  box-shadow:
    0 2px 6px rgba(0, 0, 0, 0.05),
    0 8px 20px rgba(0, 0, 0, 0.06);
}

.${c.cardClickable} {
  cursor: pointer;
}

.${c.cardClickable}:hover {
  transform: translateY(-1px);
}

.${c.cardHeader} {
  padding: 14px 16px 12px;
  border-bottom: 1px solid #f0f0f0;
  background: linear-gradient(180deg, #fafafa 0%, #ffffff 100%);
}

.${c.title} {
  font-size: 15px;
  font-weight: 600;
  line-height: 1.35;
  color: #242424;
  word-break: break-word;
}

.${c.cardBody} {
  display: flex;
  flex-direction: column;
  gap: 10px;
  padding: 12px 16px 14px;
  flex: 1 1 auto;
}

.${c.fieldRow} {
  display: grid;
  grid-template-columns: minmax(0, 38%) minmax(0, 1fr);
  gap: 4px 12px;
  align-items: baseline;
  font-size: 13px;
  line-height: 1.45;
}

.${c.fieldLabel} {
  font-size: 12px;
  font-weight: 600;
  color: #605e5c;
}

.${c.fieldValue} {
  color: #323130;
  word-break: break-word;
  min-width: 0;
}

.${c.actions} {
  display: flex;
  justify-content: flex-end;
  align-items: center;
  gap: 4px;
  padding: 10px 14px 12px;
  border-top: 1px solid #f0f0f0;
  background: #fafafa;
}

.${t.scrollWrap} .${c.grid} + .${t.empty},
.${t.scrollWrap} .${c.grid} + .${t.loading},
.${t.scrollWrap} > .${t.empty},
.${t.scrollWrap} > .${t.loading} {
  margin: 16px;
}
`.trim();
}

export const DEFAULT_CARD_LAYOUT_CSS = buildDefaultCardLayoutCss();

export function resolveCardLayoutCss(custom: string | undefined): string {
  const trimmed = (custom ?? '').trim();
  return trimmed.length > 0 ? trimmed : DEFAULT_CARD_LAYOUT_CSS;
}

export function resolveTableLayoutCss(
  slots: ITableLayoutCssSlots | undefined,
  legacyFreeform: string | undefined
): string {
  const custom = mergeCustomTableCss(slots, legacyFreeform);
  return custom.trim().length > 0 ? custom : DEFAULT_TABLE_LAYOUT_CSS;
}

export function scopeCardCssByInstance(css: string, scopeClass: string): string {
  if (!css.trim()) return '';
  const scope = `.${scopeClass}`;
  return css.replace(/(^|})\s*([^{}]+)\{/g, (match, prefix, selectorPart) => {
    const trimmed = selectorPart.trim();
    if (!trimmed || trimmed.startsWith('@') || trimmed.includes(scope)) {
      return match;
    }
    const scoped = trimmed
      .split(',')
      .map((sel: string) => `${scope} ${sel.trim()}`)
      .join(', ');
    return `${prefix} ${scoped}{`;
  });
}

export function mergeRowStyleRulesCss(rules: ITableRowStyleRule[] | undefined): string {
  if (!rules?.length) return '';
  const parts: string[] = [];
  for (let i = 0; i < rules.length; i++) {
    const r = rules[i];
    const body = (r.rowCss ?? '').trim();
    if (!body) continue;
    const token = toTableRowRuleDataToken(r.id);
    parts.push(`.${DINAMIC_SX_TABLE_CLASS.cell}[data-dinamic-rules~="${token}"] {\n${body}\n}`);
  }
  return parts.join('\n\n').trim();
}
