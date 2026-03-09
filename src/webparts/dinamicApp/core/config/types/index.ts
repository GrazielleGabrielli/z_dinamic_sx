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

export interface IPaginationConfig {
  enabled: boolean;
  pageSize: number;
  pageSizeOptions: number[];
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

export interface IListViewConfig {
  columns: IListViewColumnConfig[];
  filters: IListViewFilterConfig[];
  sort: IListViewSortConfig | null;
}

// ─── Root config ─────────────────────────────────────────────────────────────

export interface IDynamicViewConfig {
  dataSource: IDataSourceConfig;
  mode: TViewMode;
  dashboard: IDashboardConfig;
  pagination: IPaginationConfig;
  listView: IListViewConfig;
}

export interface IDynamicViewWebPartProps {
  configJson?: string;
}
