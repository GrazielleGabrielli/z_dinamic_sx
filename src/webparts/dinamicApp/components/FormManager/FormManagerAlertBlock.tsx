import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import { Stack, Text, Icon, IconButton } from '@fluentui/react';
import type { IFormFieldConfig, TFormAlertVariant } from '../../core/config/types/formManager';
import { resolveAlertVariant } from '../../core/config/types/formManager';
import { evaluateCondition } from '../../core/formManager/formRuleEngine';
import type { IDynamicContext } from '../../core/dynamicTokens/types';

export interface IFormManagerAlertBlockProps {
  alert: IFormFieldConfig;
  values: Record<string, unknown>;
  dynamicContext: IDynamicContext;
  userGroupTitles?: string[];
  fieldLabelsByName?: ReadonlyMap<string, string>;
  onConfigure?: () => void;
}

type TAlertSkin = {
  accent: string;
  border: string;
  iconColor: string;
  titleColor: string;
  bodyColor: string;
};

const DEFAULT_ICON: Record<TFormAlertVariant, string> = {
  info: 'Info',
  success: 'CheckMark',
  warning: 'Warning',
  error: 'ErrorBadge',
};

const ALERT_SKIN: Record<TFormAlertVariant, TAlertSkin> = {
  info: {
    accent: '#0078d4',
    border: '#d0e7f8',
    iconColor: '#0078d4',
    titleColor: '#323130',
    bodyColor: '#605e5c',
  },
  success: {
    accent: '#107c10',
    border: '#d5ead5',
    iconColor: '#0e700e',
    titleColor: '#323130',
    bodyColor: '#605e5c',
  },
  warning: {
    accent: '#ca5010',
    border: '#f1ddc9',
    iconColor: '#a7410f',
    titleColor: '#323130',
    bodyColor: '#605e5c',
  },
  error: {
    accent: '#a4262c',
    border: '#ebc9cb',
    iconColor: '#a4262c',
    titleColor: '#323130',
    bodyColor: '#605e5c',
  },
};

function formatAlertValue(value: unknown): string {
  if (value === null || value === undefined) return '—';
  if (typeof value === 'string') {
    const t = value.trim();
    return t || '—';
  }
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (value instanceof Date) return isNaN(value.getTime()) ? '—' : value.toLocaleString('pt-BR');
  if (Array.isArray(value)) {
    const parts = value
      .map((item) => formatAlertValue(item))
      .filter((item) => item !== '—');
    return parts.length ? parts.join('; ') : '—';
  }
  if (typeof value === 'object') {
    const rec = value as Record<string, unknown>;
    if (typeof rec.Title === 'string' && rec.Title.trim()) return rec.Title.trim();
    if (typeof rec.text === 'string' && rec.text.trim()) return rec.text.trim();
    if (typeof rec.label === 'string' && rec.label.trim()) return rec.label.trim();
    if (typeof rec.Name === 'string' && rec.Name.trim()) return rec.Name.trim();
    if (typeof rec.Id === 'number' && isFinite(rec.Id)) return String(rec.Id);
    if (typeof rec.Id === 'string' && rec.Id.trim()) return rec.Id.trim();
  }
  return String(value);
}

export const FormManagerAlertBlock: React.FC<IFormManagerAlertBlockProps> = ({
  alert,
  values,
  dynamicContext,
  userGroupTitles = [],
  fieldLabelsByName,
  onConfigure,
}) => {
  const variant = resolveAlertVariant(alert);
  const skin = ALERT_SKIN[variant];
  const title = (alert.alertTitle ?? alert.label ?? 'Alerta').trim();
  const message = (alert.alertMessage ?? alert.helpText ?? '').trim();
  const iconName = (alert.alertIconName ?? '').trim() || DEFAULT_ICON[variant];
  const alertFields = Array.isArray(alert.alertFields)
    ? alert.alertFields.map((n) => String(n).trim()).filter(Boolean)
    : [];
  const canDismiss = alert.alertDismissible === true;
  const emphasized = alert.alertEmphasized === true;
  const [dismissed, setDismissed] = useState(false);

  useEffect(() => {
    setDismissed(false);
  }, [alert.internalName, alert.alertTitle, alert.alertMessage, alert.alertVariant, alert.alertIconName]);

  const visible = useMemo(() => {
    if (!alert.alertWhen) return true;
    return evaluateCondition(alert.alertWhen, values, dynamicContext, userGroupTitles);
  }, [alert.alertWhen, values, dynamicContext, userGroupTitles]);

  if (!visible) return null;
  if (dismissed && canDismiss) return null;
  if (!title && !message && alertFields.length === 0 && onConfigure === undefined) return null;

  return (
    <Stack tokens={{ childrenGap: 3 }} styles={{ root: { marginBottom: 12, width: '100%' } }}>
      {onConfigure !== undefined ? (
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text
            variant="small"
            styles={{
              root: {
                fontWeight: 600,
                letterSpacing: '0.06em',
                textTransform: 'uppercase',
                color: '#a19f9d',
                fontSize: 11,
              },
            }}
          >
            Alerta
          </Text>
          <IconButton
            iconProps={{ iconName: 'Settings' }}
            title="Configurar bloco"
            ariaLabel="Configurar bloco"
            onClick={onConfigure}
            styles={{
              root: { width: 30, height: 30, color: '#605e5c' },
              rootHovered: { background: 'rgba(0,0,0,0.04)', color: '#323130' },
            }}
          />
        </Stack>
      ) : null}
      <div
        role={variant === 'error' ? 'alert' : 'status'}
        aria-live="polite"
        style={{
          position: 'relative',
          borderRadius: 4,
          border: `1px solid ${skin.border}`,
          background: '#ffffff',
          boxShadow: emphasized ? '0 1px 2px rgba(0,0,0,0.06)' : 'none',
          overflow: 'hidden',
          width: '100%',
        }}
      >
        <div
          style={{
            position: 'absolute',
            left: 0,
            top: 0,
            bottom: 0,
            width: 3,
            background: skin.accent,
          }}
        />
        <div style={{ padding: '10px 12px 10px 14px' }}>
          <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 8 }} styles={{ root: { minWidth: 0, width: '100%' } }}>
            <div
              style={{
                width: 20,
                height: 20,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                flexShrink: 0,
              }}
            >
              <Icon iconName={iconName} styles={{ root: { fontSize: 16, color: skin.iconColor } }} />
            </div>
            <Stack
              tokens={{ childrenGap: 0 }}
              styles={{
                root: {
                  flex: '1 1 auto',
                  minWidth: 0,
                  maxWidth: '100%',
                  paddingRight: canDismiss ? 24 : 0,
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'stretch',
                },
              }}
            >
              {title ? (
                <div
                  style={{
                    fontWeight: 600,
                    color: skin.titleColor,
                    lineHeight: 1.3,
                    fontSize: 14,
                    display: 'block',
                    margin: 0,
                    marginBottom: message ? 4 : 0,
                    whiteSpace: 'normal',
                    wordBreak: 'break-word',
                    overflowWrap: 'anywhere',
                  }}
                >
                  {title}
                </div>
              ) : null}
              {message ? (
                <div
                  style={{
                    color: skin.bodyColor,
                    lineHeight: 1.45,
                    fontSize: 12.5,
                    display: 'block',
                    margin: 0,
                    whiteSpace: 'normal',
                    wordBreak: 'break-word',
                    overflowWrap: 'anywhere',
                  }}
                >
                  {message}
                </div>
              ) : null}
              {alertFields.length > 0 ? (
                <Stack tokens={{ childrenGap: 6 }} styles={{ root: { marginTop: message ? 8 : 6 } }}>
                  {alertFields.map((fieldName) => {
                    const label = (fieldLabelsByName?.get(fieldName) ?? fieldName).trim();
                    const value = formatAlertValue(values[fieldName]);
                    return (
                      <Stack key={fieldName} horizontal tokens={{ childrenGap: 8 }} styles={{ root: { minWidth: 0 } }}>
                        <Text
                          styles={{
                            root: {
                              flex: '0 0 auto',
                              minWidth: 120,
                              fontSize: 12,
                              fontWeight: 600,
                              color: '#605e5c',
                              lineHeight: 1.4,
                            },
                          }}
                        >
                          {label}
                        </Text>
                        <Text
                          styles={{
                            root: {
                              flex: '1 1 auto',
                              minWidth: 0,
                              fontSize: 12,
                              lineHeight: 1.4,
                              color: skin.titleColor,
                              wordBreak: 'break-word',
                              overflowWrap: 'anywhere',
                            },
                          }}
                        >
                          {value}
                        </Text>
                      </Stack>
                    );
                  })}
                </Stack>
              ) : null}
            </Stack>
          </Stack>
        </div>
        {canDismiss ? (
          <IconButton
            iconProps={{ iconName: 'ChromeClose' }}
            title="Fechar"
            ariaLabel="Fechar alerta"
            onClick={() => setDismissed(true)}
            styles={{
              root: {
                position: 'absolute',
                top: 2,
                right: 2,
                width: 28,
                height: 28,
                color: '#605e5c',
              },
              rootHovered: { color: '#323130', background: 'rgba(0,0,0,0.04)' },
            }}
          />
        ) : null}
      </div>
    </Stack>
  );
};
