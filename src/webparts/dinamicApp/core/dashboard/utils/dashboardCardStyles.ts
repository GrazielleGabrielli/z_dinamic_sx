import type {
  IDashboardCardStyleConfig,
  TBorderRadius,
  TCardVariant,
  TShadow,
  TPadding,
  TTitleSize,
  TSubtitleSize,
  TValueSize,
  TFontWeight,
  TLoadingStyle,
} from '../../config/types';

const BORDER_RADIUS_MAP: Record<TBorderRadius, string> = {
  none: '0',
  sm: '0.25rem',
  md: '0.375rem',
  lg: '0.5rem',
  xl: '0.75rem',
  full: '9999px',
};

const PADDING_MAP: Record<TPadding, string> = {
  sm: '0.5rem 0.75rem',
  md: '1rem 1.25rem',
  lg: '1.25rem 1.5rem',
};

const SHADOW_MAP: Record<TShadow, string> = {
  none: 'none',
  sm: '0 1px 2px 0 rgb(0 0 0 / 0.05)',
  md: '0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1)',
  lg: '0 10px 15px -3px rgb(0 0 0 / 0.1), 0 4px 6px -4px rgb(0 0 0 / 0.1)',
};

const TITLE_SIZE_MAP: Record<TTitleSize, string> = {
  xs: '0.75rem',
  sm: '0.875rem',
  md: '1rem',
  lg: '1.125rem',
};

const SUBTITLE_SIZE_MAP: Record<TSubtitleSize, string> = {
  xs: '0.75rem',
  sm: '0.8125rem',
  md: '0.875rem',
};

const VALUE_SIZE_MAP: Record<TValueSize, string> = {
  lg: '1.25rem',
  xl: '1.5rem',
  '2xl': '1.875rem',
  '3xl': '2.25rem',
};

const FONT_WEIGHT_MAP: Record<TFontWeight, string> = {
  normal: '400',
  medium: '500',
  semibold: '600',
  bold: '700',
};

export function getDefaultDashboardCardStyle(): IDashboardCardStyleConfig {
  return {
    variant: 'default',
    borderRadius: 'lg',
    padding: 'md',
    shadow: 'sm',
    border: true,
    titleSize: 'sm',
    subtitleSize: 'xs',
    valueSize: '2xl',
    titleWeight: 'semibold',
    valueWeight: 'bold',
    align: 'left',
    showIcon: false,
    iconPosition: 'left',
    showSubtitle: true,
    showValue: true,
    loadingStyle: 'skeleton',
  };
}

export function mergeWithDefaultStyle(
  partial?: Partial<IDashboardCardStyleConfig> | null
): IDashboardCardStyleConfig {
  const def = getDefaultDashboardCardStyle();
  if (!partial || typeof partial !== 'object') return { ...def };
  return { ...def, ...partial };
}

function getVariantStyles(variant: TCardVariant): { background?: string; border?: string } {
  switch (variant) {
    case 'outlined':
      return { background: 'transparent', border: '1px solid' };
    case 'soft':
      return { background: 'var(--card-bg-soft, #f8fafc)' };
    case 'solid':
      return { background: 'var(--card-bg-solid, #0f172a)' };
    default:
      return { background: 'var(--card-bg-default, #ffffff)' };
  }
}

export function getCardContainerClasses(style: IDashboardCardStyleConfig): string[] {
  const classes: string[] = [];
  if (style.align === 'center') classes.push('text-center');
  if (style.align === 'right') classes.push('text-right');
  return classes;
}

export function getCardInlineStyles(style: IDashboardCardStyleConfig): Record<string, string> {
  const variantStyles = getVariantStyles(style.variant);
  const borderValue = style.border
    ? style.borderColor
      ? `1px solid ${style.borderColor}`
      : variantStyles.border
        ? `${variantStyles.border} var(--card-border, #e2e8f0)`
        : '1px solid var(--card-border, #e2e8f0)'
    : 'none';
  return {
    borderRadius: BORDER_RADIUS_MAP[style.borderRadius],
    padding: PADDING_MAP[style.padding],
    boxShadow: SHADOW_MAP[style.shadow],
    border: borderValue,
    backgroundColor: style.backgroundColor ?? variantStyles.background ?? '#ffffff',
    textAlign: style.align,
  };
}

export interface ICardTextStyles {
  title: Record<string, string>;
  subtitle: Record<string, string>;
  value: Record<string, string>;
}

export function getCardTextStyles(style: IDashboardCardStyleConfig): ICardTextStyles {
  return {
    title: {
      fontSize: TITLE_SIZE_MAP[style.titleSize],
      fontWeight: FONT_WEIGHT_MAP[style.titleWeight],
      color: style.titleColor ?? 'inherit',
    },
    subtitle: {
      fontSize: SUBTITLE_SIZE_MAP[style.subtitleSize],
      color: style.subtitleColor ?? 'inherit',
    },
    value: {
      fontSize: VALUE_SIZE_MAP[style.valueSize],
      fontWeight: FONT_WEIGHT_MAP[style.valueWeight],
      color: style.valueColor ?? 'inherit',
    },
  };
}

export function getValueDisplayColor(
  style: IDashboardCardStyleConfig,
  value: number | undefined
): string | undefined {
  if (value === undefined) return undefined;
  if (style.highlightNegative && value < 0) return 'var(--negative, #dc2626)';
  if (style.highlightZero && value === 0) return 'var(--zero, #64748b)';
  return style.valueColor;
}

export type { TLoadingStyle };
