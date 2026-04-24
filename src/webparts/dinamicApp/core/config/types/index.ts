export type TViewMode = 'list' | 'projectManagement' | 'formManager';
export type TSourceKind = 'list' | 'library';

// ─── Data source ────────────────────────────────────────────────────────────

export interface IDataSourceConfig {
  kind: TSourceKind;
  title: string;
}

// ─── Dashboard ───────────────────────────────────────────────────────────────

export type TAggregateType = 'count' | 'sum';
export type TFilterOperator = 'eq' | 'ne' | 'gt' | 'lt' | 'ge' | 'le' | 'contains';

export interface IDashboardCardFilter {
  field: string;
  operator: TFilterOperator;
  value: string;
}

export type TCardVariant = 'default' | 'outlined' | 'soft' | 'solid';
export type TBorderRadius = 'none' | 'sm' | 'md' | 'lg' | 'xl' | 'full';
export type TPadding = 'sm' | 'md' | 'lg';
export type TShadow = 'none' | 'sm' | 'md' | 'lg';
export type TTitleSize = 'xs' | 'sm' | 'md' | 'lg';
export type TSubtitleSize = 'xs' | 'sm' | 'md';
export type TValueSize = 'lg' | 'xl' | '2xl' | '3xl';
export type TFontWeight = 'normal' | 'medium' | 'semibold' | 'bold';
export type TAlign = 'left' | 'center' | 'right';
export type TIconPosition = 'left' | 'top' | 'right';
export type TLoadingStyle = 'skeleton' | 'spinner' | 'text';

export interface IDashboardCardStyleConfig {
  variant: TCardVariant;
  borderRadius: TBorderRadius;
  padding: TPadding;
  shadow: TShadow;
  border: boolean;
  backgroundColor?: string;
  borderColor?: string;
  titleColor?: string;
  subtitleColor?: string;
  valueColor?: string;
  iconColor?: string;
  titleSize: TTitleSize;
  subtitleSize: TSubtitleSize;
  valueSize: TValueSize;
  titleWeight: TFontWeight;
  valueWeight: TFontWeight;
  align: TAlign;
  showIcon: boolean;
  iconName?: string;
  iconPosition: TIconPosition;
  showSubtitle: boolean;
  showValue: boolean;
  highlightNegative?: boolean;
  highlightZero?: boolean;
  loadingStyle: TLoadingStyle;
}

export interface IDashboardCardConfig {
  id: string;
  title: string;
  aggregate: TAggregateType;
  field?: string;
  /** Para campo lookup: campo da lista de destino (ex: Title). Gera $expand e select campo/expandField */
  expandField?: string;
  filter?: IDashboardCardFilter;
  filters?: IDashboardCardFilter[];
  subtitle?: string;
  emptyValueText?: string;
  errorText?: string;
  loadingText?: string;
  style?: IDashboardCardStyleConfig;
}

export type TDashboardType = 'cards' | 'charts';
export type TChartType = 'bar' | 'line' | 'area' | 'pie' | 'donut';

export interface IChartSeriesConfig {
  id: string;
  label: string;
  aggregate: TAggregateType;
  field?: string;
  expandField?: string;
  filter?: IDashboardCardFilter;
  filters?: IDashboardCardFilter[];
  color?: string;
}

export interface IDashboardConfig {
  enabled: boolean;
  dashboardType: TDashboardType;
  cardsCount: number;
  cards: IDashboardCardConfig[];
  chartType?: TChartType;
  chartSeries?: IChartSeriesConfig[];
}

// ─── Pagination ──────────────────────────────────────────────────────────────

export type TPaginationLayout = 'buttons' | 'numbered' | 'compact' | 'paged';

export interface IPaginationConfig {
  enabled: boolean;
  pageSize: number;
  pageSizeOptions: number[];
  /** Layout da paginação exibida após a tabela */
  layout?: TPaginationLayout;
}

// ─── List view ───────────────────────────────────────────────────────────────

export interface IListViewColumnConfig {
  field: string;
  label?: string;
  width?: number;
  /** Para campo lookup: campo da lista de destino a exibir (ex: Title). Gera $expand e select campo/expandField */
  expandField?: string;
}

export interface IListViewFilterConfig {
  field: string;
  operator: TFilterOperator;
  value: string;
}

export interface IListViewSortConfig {
  field: string;
  ascending: boolean;
}

export interface IListViewModeConfig {
  id: string;
  label: string;
  filters: IListViewFilterConfig[];
}

/** Modo inicial da lista quando Tabela/Cards está ativo. */
export type TListViewDisplayMode = 'table' | 'cards';

export type TTableCssSlot =
  | 'viewRoot'
  | 'toolbar'
  | 'scrollWrap'
  | 'table'
  | 'thead'
  | 'headerRow'
  | 'headerCell'
  | 'headerCellInner'
  | 'headerFilterTrigger'
  | 'body'
  | 'row'
  | 'cell'
  | 'empty'
  | 'loading'
  | 'error'
  | 'pagination';

export type ITableLayoutCssSlots = Partial<Record<TTableCssSlot, string>>;

export type TTableRowRuleOperator =
  | 'eq'
  | 'ne'
  | 'contains'
  | 'startsWith'
  | 'endsWith'
  | 'empty'
  | 'notEmpty';

export interface ITableRowStyleRule {
  id: string;
  /** Nome interno do campo (ex.: Title). */
  field: string;
  operator: TTableRowRuleOperator;
  value: string;
  /** Declarações CSS na linha (<tr>) quando a condição for verdadeira. */
  rowCss: string;
}

export type TListRowActionIconPreset = 'view' | 'edit' | 'link' | 'custom';

/** `icon` = só o botão; `wholeRow` = linha da tabela ou card inteiro também abre a URL desta ação. */
export type TListRowActionScope = 'icon' | 'wholeRow';

export interface IListRowActionConfig {
  id: string;
  title: string;
  iconPreset: TListRowActionIconPreset;
  /** Quando iconPreset é custom, nome do ícone Fluent (ex.: Mail, Share). */
  customIconName?: string;
  /** URL com `{Campo}` ou `{Lookup/Campo}` e tokens dinâmicos [me], [siteurl], etc. */
  urlTemplate: string;
  openInNewTab?: boolean;
  scope: TListRowActionScope;
}

export interface IListViewConfig {
  columns: IListViewColumnConfig[];
  filters: IListViewFilterConfig[];
  sort: IListViewSortConfig | null;
  viewModes?: IListViewModeConfig[];
  activeViewModeId?: string;
  pdfExportEnabled?: boolean;
  /** Quando true, a lista exibe alternância Tabela / Cards na barra de ferramentas. */
  listCardViewEnabled?: boolean;
  /** Com `listCardViewEnabled`, modo ao carregar (omitido = tabela). */
  listDefaultDisplayMode?: TListViewDisplayMode;
  /** Declarações CSS por região da tabela (aba Layout); cada bloco é aplicado à classe correspondente. */
  customTableCssSlots?: ITableLayoutCssSlots;
  /** CSS livre (regras completas, seletores combinados, [data-field], etc.). */
  customTableCss?: string;
  /** Estilo condicional por linha conforme valor de coluna (aba Layout → Regras). */
  tableRowStyleRules?: ITableRowStyleRule[];
  /** Ações por item (ícones e/ou clique na linha/card). Configurado na aba Ações do painel. */
  listRowActions?: IListRowActionConfig[];
}

// ─── Project management ─────────────────────────────────────────────────────

export interface IProjectManagementRuleConfig {
  id: string;
  field: string;
  value: string;
}

export interface IProjectManagementColumnConfig {
  id: string;
  title: string;
  rules: IProjectManagementRuleConfig[];
}

export interface IProjectManagementConfig {
  columns: IProjectManagementColumnConfig[];
}

// ─── PDF template ───────────────────────────────────────────────────────────

export type TPdfElementType = 'text' | 'image' | 'rect' | 'line';
export type TPdfElementScope = 'fixed' | 'dynamic';

export interface IPdfTemplateElement {
  id: string;
  type: TPdfElementType;
  scope?: TPdfElementScope;
  x: number;
  y: number;
  width?: number;
  height?: number;
  content?: string;
  fontSize?: number;
  fontWeight?: 'normal' | 'bold';
  color?: string;
  imageUrl?: string;
}

export interface IPdfTemplateSection {
  height?: number;
  elements: IPdfTemplateElement[];
}

export type TPdfLayoutMode = 'onePerPage' | 'allOnOnePage' | 'breakWhenFull';

export type TPdfPageFormat =
  | 'A0'
  | 'A1'
  | 'A2'
  | 'A3'
  | 'A4'
  | 'A5'
  | 'A6'
  | 'B4'
  | 'B5'
  | 'Letter'
  | 'Legal'
  | 'Tabloid'
  | 'CreditCard';

export interface IPdfTemplateConfig {
  pageFormat: TPdfPageFormat;
  orientation: 'portrait' | 'landscape';
  layoutMode?: TPdfLayoutMode;
  bodyBlockHeightMm?: number;
  fixedBlockHeightMm?: number;
  header?: IPdfTemplateSection;
  footer?: IPdfTemplateSection;
  body: IPdfTemplateSection;
}

// ─── Root config ─────────────────────────────────────────────────────────────

/** Layout de colunas por seção (estilo página moderna). */
export type TListPageSectionLayout = 'one' | 'two' | 'three' | 'oneThirdLeft' | 'oneThirdRight';

export type TListPageBlockType =
  | 'dashboard'
  | 'list'
  | 'banner'
  | 'editor'
  | 'sectionTitle'
  | 'alert'
  | 'buttons';

/** Ação de um botão no bloco «Botões» (modo lista). */
export type TListPageButtonActionKind = 'redirect' | 'reload';

export interface IListPageButtonItemConfig {
  id: string;
  label: string;
  actionKind: TListPageButtonActionKind;
  /** Obrigatório quando `actionKind` é `redirect`. */
  url?: string;
  openInNewTab?: boolean;
}

export interface IListPageButtonsBlockConfig {
  items: IListPageButtonItemConfig[];
}

export type TListPageBannerContentAlign = 'left' | 'center' | 'right';

export type TListPageSectionTitleSize = 'sm' | 'md' | 'lg';

export interface IListPageSectionTitleBlockConfig {
  title: string;
  subtitle: string;
  iconName: string;
  align: TListPageBannerContentAlign;
  showDivider: boolean;
  size: TListPageSectionTitleSize;
  marginTopPx: number;
  marginBottomPx: number;
}

export type TListPageAlertVariant = 'info' | 'success' | 'warning' | 'error';

/** Comparar o número de itens da lista (com filtro OData) com um valor. */
export type TListPageAlertCountOp = 'eq' | 'ne' | 'gt' | 'ge' | 'lt' | 'le';

/** Operador OData no campo escolhido para a contagem (separado do `countOp` que compara o total). */
export type TListPageAlertCountFilterFieldOp =
  | 'eq'
  | 'ne'
  | 'gt'
  | 'ge'
  | 'lt'
  | 'le'
  | 'contains';

/**
 * Regra por contagem na lista da vista. A primeira regra cuja contagem corresponder define o aspeto
 * (sobrepondo título, mensagem, tipo, ícone, etc. ao «padrão»).
 */
export interface IListPageAlertCountRule {
  id: string;
  /** Filtro OData (ex.: `Status eq 'Aberto'`). Vazio = contar todos os itens (até 5000). */
  odataFilter?: string;
  /** Construtor visual: campo interno da lista; ausente = sem filtro por campo (ou só OData manual). */
  countFilterField?: string;
  countFilterFieldOp?: TListPageAlertCountFilterFieldOp;
  /** Valor do filtro no campo (texto, número, Id de lookup, etc.). */
  countFilterValue?: string;
  /** Quando true, mostrar e editar «odataFilter» em texto livre em vez do construtor. */
  countFilterUseManualOdata?: boolean;
  countOp: TListPageAlertCountOp;
  count: number;
  title?: string;
  message?: string;
  variant?: TListPageAlertVariant;
  iconName?: string;
  dismissible?: boolean;
  emphasized?: boolean;
  linkUrl?: string;
  linkText?: string;
}

export interface IListPageAlertBlockConfig {
  title: string;
  message: string;
  variant: TListPageAlertVariant;
  iconName: string;
  dismissible: boolean;
  emphasized: boolean;
  linkUrl: string;
  linkText: string;
  /** Avaliadas por ordem; a primeira que coincidir substitui o aspeto padrão acima. */
  countRules?: IListPageAlertCountRule[];
}

export interface IListPageBannerBlockConfig {
  imageUrl: string;
  title: string;
  subtitle: string;
  linkUrl: string;
  openInNewTab: boolean;
  imageAlt: string;
  contentAlign: TListPageBannerContentAlign;
  heightPx: number;
  /** 0–1 escurecimento sobre a imagem */
  overlayOpacity: number;
  showButton: boolean;
  buttonText: string;
}

export interface IListPageRichEditorBlockConfig {
  title: string;
  /** HTML armazenado; filtrado na exibição conforme permissões */
  html: string;
  placeholder: string;
  minHeightPx: number;
  readOnly: boolean;
  allowImages: boolean;
  allowLinks: boolean;
  allowTables: boolean;
  allowLists: boolean;
  allowHeaders: boolean;
  allowVideoEmbed: boolean;
}

/** Lista filha usada por um bloco do layout (dados OData / metadados dessa lista). */
export interface IListPageLinkedListBinding {
  listTitle: string;
  /** Lookup na lista filha que aponta para a lista principal do app. */
  parentLookupFieldInternalName: string;
}

export interface IListPageBlock {
  id: string;
  type: TListPageBlockType;
  /** Só em `dashboard`. Se ausente e houver um único bloco dashboard, usa `IDynamicViewConfig.dashboard`. */
  dashboard?: IDashboardConfig;
  banner?: IListPageBannerBlockConfig;
  editor?: IListPageRichEditorBlockConfig;
  sectionTitle?: IListPageSectionTitleBlockConfig;
  alert?: IListPageAlertBlockConfig;
  buttons?: IListPageButtonsBlockConfig;
  /** Blocos `dashboard` / `list` / `alert`: dados desta lista (com lookup para a principal). */
  linkedListBinding?: IListPageLinkedListBinding;
}

export interface IListPageSection {
  id: string;
  layout: TListPageSectionLayout;
  /** Uma entrada por coluna; cada coluna é uma pilha de blocos. */
  columns: IListPageBlock[][];
}

export interface IListPageLayoutConfig {
  sections: IListPageSection[];
  /**
   * Espaçamento interno da área do layout (CSS padding), ex.: «16px 24px» (vertical horizontal).
   * Apenas valores «Npx» separados por espaços (1 a 4 valores).
   */
  contentPadding?: string;
}

export type {
  IFormManagerConfig,
  IFormManagerActionLogConfig,
  IFormStepNavigationConfig,
  IFormCustomButtonConfig,
  TFormCustomButtonFinishAfterRun,
  TFormButtonAction,
  TFormCustomButtonBehavior,
  TFormCustomButtonOperation,
  TFormRule,
  TFormManagerFormMode,
  TFormStepLayoutKind,
  TFormStepNavButtonsKind,
  TFormDataLoadingUiKind,
  TFormSubmitLoadingUiKind,
  TFormAttachmentUploadLayoutKind,
  TFormAttachmentFilePreviewKind,
  TFormAttachmentStorageKind,
  IFormManagerAttachmentLibraryConfig,
  TFormHistoryPresentationKind,
  TFormHistoryLayoutKind,
  TFormHistoryButtonKind,
  TFormHistoryIntegratedClickBehavior,
  TFormCustomButtonPaletteSlot,
  TFormRootWidthMode,
  TFormRootHorizontalAlign,
  FORM_ATTACHMENTS_FIELD_INTERNAL,
  FORM_OCULTOS_STEP_ID,
  FORM_FIXOS_STEP_ID,
  FORM_BUILTIN_HISTORY_BUTTON_ID,
} from './formManager';
export type {
  IFormFieldConfig,
  IFormSectionConfig,
  IFormStepConfig,
  TFormConditionNode,
  TFormFieldTextValueTransform,
  TFormFieldTextInputMaskKind,
  TFormSubmitKind,
} from './formManager';

export interface IModeConfigSnapshot {
  listView?: IListViewConfig;
  projectManagement?: IProjectManagementConfig;
  formManager?: import('./formManager').IFormManagerConfig;
  dashboard?: IDashboardConfig;
  pagination?: IPaginationConfig;
  listPageLayout?: IListPageLayoutConfig;
  pdfTemplate?: IPdfTemplateConfig;
  tableConfig?: import('../../table').ITableConfig;
}

export interface IConfigMemory {
  /** Chave: `${kind}::${title}` (title sem espaços nas extremidades). */
  bySource: Record<string, Partial<Record<TViewMode, IModeConfigSnapshot>>>;
}

export interface IDynamicViewConfig {
  dataSource: IDataSourceConfig;
  mode: TViewMode;
  dashboard: IDashboardConfig;
  pagination: IPaginationConfig;
  listView: IListViewConfig;
  projectManagement?: IProjectManagementConfig;
  /** Modo formulário + gestor: layout, campos e regras dinâmicas. */
  formManager?: import('./formManager').IFormManagerConfig;
  /** Config da tabela dinâmica (modo list). Quando presente, DataTable + TableEngine são usados. */
  tableConfig?: import('../../table').ITableConfig;
  pdfTemplate?: IPdfTemplateConfig;
  /**
   * Modo lista: seções com colunas e blocos (dashboard / tabela).
   * Se ausente, usa o layout legado (dashboard acima + título + tabela).
   */
  listPageLayout?: IListPageLayoutConfig;
  /** Memória por lista/biblioteca e modo (alternância no assistente). */
  configMemory?: IConfigMemory;
}

export interface IDynamicViewWebPartProps {
  configJson?: string;
}
