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
} as const;

export const DINAMIC_SX_CARD_CLASS = {
  grid: 'dinamicSxCardGrid',
  card: 'dinamicSxCard',
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

export function scopeCardCssByInstance(css: string, scopeClass: string): string {
  if (!css.trim()) return '';
  return css.replace(/\.dinamicSxCard/g, `.${scopeClass} .dinamicSxCard`);
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
