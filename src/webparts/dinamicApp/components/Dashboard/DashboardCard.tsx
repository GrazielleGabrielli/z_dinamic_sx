import * as React from 'react';
import { Spinner, SpinnerSize, Icon } from '@fluentui/react';
import { IDashboardCardResult } from '../../core/dashboard/types';
import { IDashboardCardConfig } from '../../core/config/types';
import {
  mergeWithDefaultStyle,
  getCardContainerClasses,
  getCardInlineStyles,
  getCardTextStyles,
  getValueDisplayColor,
} from '../../core/dashboard/utils';

interface IDashboardCardProps {
  result: IDashboardCardResult;
  cardConfig?: IDashboardCardConfig;
  selected?: boolean;
  onActivate?: () => void;
}

function formatValue(value: number): string {
  return value.toLocaleString('pt-BR', { maximumFractionDigits: 2 });
}

export const DashboardCard: React.FC<IDashboardCardProps> = ({ result, cardConfig, selected, onActivate }) => {
  const style = mergeWithDefaultStyle(cardConfig?.style);
  const containerStyles = getCardInlineStyles(style);
  const textStyles = getCardTextStyles(style);
  const containerClasses = getCardContainerClasses(style);

  const title = cardConfig?.title ?? result.title;
  const subtitle = cardConfig?.subtitle;
  const emptyValueText = cardConfig?.emptyValueText ?? '—';
  const errorText = cardConfig?.errorText ?? 'Erro ao carregar';
  const loadingText = cardConfig?.loadingText ?? 'Carregando...';

  const showSubtitle = style.showSubtitle && subtitle !== undefined && subtitle !== '';
  const showValue = style.showValue;

  const valueColor = getValueDisplayColor(style, result.value);
  const displayValue =
    result.status === 'ready' && result.value !== undefined && result.value !== null
      ? formatValue(result.value)
      : result.status === 'ready' && (result.value === undefined || result.value === null)
        ? emptyValueText
        : null;

  const className = ['dashboard-card', ...containerClasses].filter(Boolean).join(' ');
  const clickable = Boolean(onActivate) && result.status === 'ready';

  return (
    <div
      className={className}
      role={clickable ? 'button' : undefined}
      tabIndex={clickable ? 0 : undefined}
      aria-pressed={clickable ? selected : undefined}
      onClick={clickable ? () => onActivate?.() : undefined}
      onKeyDown={
        clickable
          ? (e) => {
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                onActivate?.();
              }
            }
          : undefined
      }
      style={{
        ...containerStyles,
        flex: '1 1 180px',
        minWidth: 160,
        maxWidth: 280,
        cursor: clickable ? 'pointer' : undefined,
        outline: selected ? '2px solid #0078d4' : undefined,
        outlineOffset: selected ? 2 : undefined,
      }}
    >
      {result.status === 'loading' && (
        <>
          {style.loadingStyle === 'spinner' && (
            <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', minHeight: 72 }}>
              <Spinner size={SpinnerSize.medium} />
            </div>
          )}
          {style.loadingStyle === 'text' && (
            <div style={{ ...textStyles.subtitle, padding: '0.5rem 0' }}>{loadingText}</div>
          )}
          {style.loadingStyle === 'skeleton' && (
            <div style={{ minHeight: 72 }}>
              <div style={{ ...textStyles.title, opacity: 0.6, marginBottom: 4 }}>{title}</div>
              <div
                style={{
                  height: 28,
                  backgroundColor: 'rgba(0,0,0,0.08)',
                  borderRadius: 4,
                  maxWidth: 80,
                }}
              />
            </div>
          )}
        </>
      )}

      {result.status === 'ready' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          {style.showIcon && style.iconName && style.iconPosition === 'top' && (
            <Icon iconName={style.iconName} styles={{ root: { color: style.iconColor ?? 'inherit', fontSize: '1.25rem' } }} />
          )}
          <div style={{ display: 'flex', alignItems: 'flex-start', gap: 8 }}>
            {style.showIcon && style.iconName && style.iconPosition === 'left' && (
              <Icon iconName={style.iconName} styles={{ root: { color: style.iconColor ?? 'inherit', fontSize: '1.125rem' } }} />
            )}
            <div style={{ flex: 1 }}>
              <div style={textStyles.title}>{title}</div>
              {showSubtitle && <div style={{ ...textStyles.subtitle, marginTop: 2 }}>{subtitle}</div>}
              {showValue && (
                <div style={{ ...textStyles.value, color: valueColor ?? textStyles.value.color, marginTop: 4 }}>
                  {displayValue ?? emptyValueText}
                </div>
              )}
            </div>
            {style.showIcon && style.iconName && style.iconPosition === 'right' && (
              <Icon iconName={style.iconName} styles={{ root: { color: style.iconColor ?? 'inherit', fontSize: '1.125rem' } }} />
            )}
          </div>
        </div>
      )}

      {result.status === 'error' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          <div style={textStyles.title}>{title}</div>
          <div style={{ ...textStyles.subtitle, color: 'var(--error, #dc2626)' }}>
            {result.error ?? errorText}
          </div>
        </div>
      )}
    </div>
  );
};
