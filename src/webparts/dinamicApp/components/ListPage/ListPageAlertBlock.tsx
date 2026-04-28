import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import { Link, Stack, Text, IconButton, Spinner, SpinnerSize, Icon } from '@fluentui/react';
import type { IListPageAlertBlockConfig, TListPageAlertVariant } from '../../core/config/types';
import {
  defaultAlertConfig,
  listAlertCountMatches,
  mergeAlertWithCountRule,
} from '../../core/listPage/listPageBlockConfigUtils';
import { ItemsService } from '../../../../services';

export interface IListPageAlertBlockProps {
  alert?: IListPageAlertBlockConfig;
  /** Lista da vista (contagem OData). */
  listTitle?: string;
  onConfigure?: () => void;
}

const DEFAULT_ICON: Record<TListPageAlertVariant, string> = {
  info: 'Info',
  success: 'CheckMark',
  warning: 'Warning',
  error: 'ErrorBadge',
};

type TVariantSkin = {
  accent: string;
  surface: string;
  border: string;
  iconBg: string;
  iconColor: string;
  titleColor: string;
  bodyColor: string;
};

const VARIANT_SKIN: Record<TListPageAlertVariant, TVariantSkin> = {
  info: {
    accent: '#0078d4',
    surface: 'linear-gradient(135deg, #f6f9fc 0%, #f0f5fa 100%)',
    border: 'rgba(0, 120, 212, 0.18)',
    iconBg: 'rgba(0, 120, 212, 0.1)',
    iconColor: '#0078d4',
    titleColor: '#201f1e',
    bodyColor: '#605e5c',
  },
  success: {
    accent: '#107c10',
    surface: 'linear-gradient(135deg, #f4faf4 0%, #edf7ed 100%)',
    border: 'rgba(16, 124, 16, 0.22)',
    iconBg: 'rgba(16, 124, 16, 0.12)',
    iconColor: '#0e700e',
    titleColor: '#201f1e',
    bodyColor: '#605e5c',
  },
  warning: {
    accent: '#ca5010',
    surface: 'linear-gradient(135deg, #fffbf5 0%, #fff4e6 100%)',
    border: 'rgba(202, 80, 16, 0.25)',
    iconBg: 'rgba(202, 80, 16, 0.12)',
    iconColor: '#a7410f',
    titleColor: '#323130',
    bodyColor: '#605e5c',
  },
  error: {
    accent: '#a4262c',
    surface: 'linear-gradient(135deg, #fdf6f6 0%, #fce8e8 100%)',
    border: 'rgba(164, 38, 44, 0.22)',
    iconBg: 'rgba(164, 38, 44, 0.1)',
    iconColor: '#a4262c',
    titleColor: '#323130',
    bodyColor: '#605e5c',
  },
};

export const ListPageAlertBlock: React.FC<IListPageAlertBlockProps> = ({
  alert: raw,
  listTitle = '',
  onConfigure,
}) => {
  const base = useMemo(() => {
    if (!raw) return defaultAlertConfig();
    return {
      ...raw,
      countRules: raw.countRules?.map((r) => ({ ...r })),
    };
  }, [
    raw?.title,
    raw?.message,
    raw?.variant,
    raw?.iconName,
    raw?.dismissible,
    raw?.emphasized,
    raw?.linkUrl,
    raw?.linkText,
    JSON.stringify(raw?.countRules ?? null),
  ]);
  const itemsService = useMemo(() => new ItemsService(), []);
  const [effective, setEffective] = useState<IListPageAlertBlockConfig>(base);
  const [countLoading, setCountLoading] = useState(false);
  const [countErr, setCountErr] = useState<string | undefined>(undefined);

  useEffect(() => {
    const rules = base.countRules ?? [];
    if (!rules.length) {
      setEffective(base);
      setCountLoading(false);
      setCountErr(undefined);
      return;
    }
    const lt = listTitle.trim();
    if (!lt) {
      setEffective(base);
      setCountLoading(false);
      setCountErr(undefined);
      return;
    }
    let cancel = false;
    setCountLoading(true);
    setCountErr(undefined);
    void (async () => {
      try {
        for (let i = 0; i < rules.length; i++) {
          const rule = rules[i];
          const cnt = await itemsService.countItems(lt, rule.odataFilter);
          if (cancel) return;
          if (listAlertCountMatches(cnt, rule.countOp, rule.count)) {
            if (!cancel) setEffective(mergeAlertWithCountRule(base, rule));
            if (!cancel) setCountLoading(false);
            return;
          }
        }
        if (!cancel) setEffective(base);
      } catch (e) {
        if (!cancel) {
          setEffective(base);
          setCountErr(e instanceof Error ? e.message : String(e));
        }
      } finally {
        if (!cancel) setCountLoading(false);
      }
    })();
    return (): void => {
      cancel = true;
    };
  }, [base, itemsService, listTitle]);

  const c = effective;
  const [dismissed, setDismissed] = useState(false);
  useEffect(() => {
    setDismissed(false);
  }, [c.title, c.message, c.variant, c.iconName]);

  const iconOverride = c.iconName.trim();
  const hasLink = Boolean(c.linkUrl.trim() && c.linkText.trim());
  const hasBody = Boolean(c.title.trim() || c.message.trim() || hasLink);
  if (!hasBody && onConfigure === undefined) {
    return null;
  }

  const toolbar =
    onConfigure !== undefined ? (
      <Stack
        horizontal
        horizontalAlign="space-between"
        verticalAlign="center"
        styles={{
          root: {
            marginBottom: 10,
            padding: '2px 4px 2px 2px',
          },
        }}
      >
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
            root: { width: 32, height: 32, color: '#0078d4' },
            icon: { fontSize: 16 },
          }}
        />
      </Stack>
    ) : null;

  if (dismissed && c.dismissible) {
    return (
      <Stack tokens={{ childrenGap: 8 }}>
        {toolbar}
        {countLoading && (
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
            <Spinner size={SpinnerSize.small} />
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              A avaliar regras de contagem…
            </Text>
          </Stack>
        )}
      </Stack>
    );
  }

  const skin = VARIANT_SKIN[c.variant];
  const iconName = iconOverride || DEFAULT_ICON[c.variant];
  const shadow = c.emphasized
    ? '0 4px 14px rgba(0, 0, 0, 0.08), 0 0 0 1px rgba(0,0,0,0.04)'
    : '0 2px 8px rgba(0, 0, 0, 0.04)';

  const inner = (
    <Stack
      horizontal
      verticalAlign="center"
      tokens={{ childrenGap: 14 }}
      styles={{ root: { alignItems: 'center', minWidth: 0 } }}
    >
      <div
        style={{
          width: 44,
          height: 44,
          borderRadius: 12,
          background: skin.iconBg,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          flexShrink: 0,
        }}
      >
        <Icon iconName={iconName} styles={{ root: { fontSize: 22, color: skin.iconColor } }} />
      </div>
      <Stack
        tokens={{ childrenGap: 16 }}
        styles={{
          root: {
            flex: '1 1 auto',
            minWidth: 0,
            maxWidth: '100%',
            paddingRight: c.dismissible ? 36 : 0,
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'stretch',
          },
        }}
      >
        {countErr ? (
          <Text block variant="small" styles={{ root: { color: '#a4262c', lineHeight: 1.45 } }}>
            Contagem: {countErr}
          </Text>
        ) : null}
        {c.title.trim() ? (
          <Text
            block
            variant="mediumPlus"
            styles={{
              root: {
                fontWeight: 700,
                color: skin.titleColor,
                letterSpacing: '-0.01em',
                lineHeight: 1.4,
                wordWrap: 'break-word',
                overflowWrap: 'break-word',
              },
            }}
          >
            {c.title}
          </Text>
        ) : null}
        {c.message.trim() ? (
          <Text
            block
            variant="small"
            styles={{
              root: {
                color: skin.bodyColor,
                lineHeight: 1.55,
                fontSize: 13,
                wordWrap: 'break-word',
                overflowWrap: 'break-word',
              },
            }}
          >
            {c.message}
          </Text>
        ) : (
          onConfigure !== undefined && (
            <Text block variant="small" styles={{ root: { color: '#a19f9d', fontStyle: 'italic', fontSize: 13 } }}>
              Sem mensagem — configure o bloco
            </Text>
          )
        )}
        {hasLink ? (
          <Link
            href={c.linkUrl.trim()}
            target="_blank"
            rel="noopener noreferrer"
            styles={{
              root: {
                marginTop: 2,
                fontWeight: 600,
                fontSize: 13,
                display: 'inline-block',
                wordWrap: 'break-word',
              },
            }}
          >
            {c.linkText.trim()} →
          </Link>
        ) : null}
      </Stack>
    </Stack>
  );

  return (
    <Stack className="dinamicSxAlert" tokens={{ childrenGap: 10 }}>
      {toolbar}
      {countLoading && !dismissed ? (
        <Stack
          horizontal
          verticalAlign="center"
          tokens={{ childrenGap: 10 }}
          styles={{
            root: {
              padding: '10px 14px',
              borderRadius: 10,
              background: '#faf9f8',
              border: '1px solid #edebe9',
            },
          }}
        >
          <Spinner size={SpinnerSize.small} />
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            A avaliar regras de contagem…
          </Text>
        </Stack>
      ) : null}
      <div
        role={c.variant === 'error' ? 'alert' : 'status'}
        aria-live="polite"
        style={{
          position: 'relative',
          borderRadius: 14,
          border: `1px solid ${skin.border}`,
          background: skin.surface,
          boxShadow: shadow,
          overflow: 'hidden',
        }}
      >
        <div
          style={{
            position: 'absolute',
            left: 0,
            top: 0,
            bottom: 0,
            width: 4,
            background: skin.accent,
            borderRadius: '14px 0 0 14px',
          }}
        />
        <div style={{ padding: '18px 20px 18px 22px' }}>{inner}</div>
        {c.dismissible ? (
          <IconButton
            iconProps={{ iconName: 'ChromeClose' }}
            title="Fechar"
            ariaLabel="Fechar alerta"
            onClick={() => setDismissed(true)}
            styles={{
              root: {
                position: 'absolute',
                top: 6,
                right: 4,
                width: 32,
                height: 32,
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
