import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import { Link, MessageBar, MessageBarType, Stack, Text, ActionButton, Spinner } from '@fluentui/react';
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

function messageBarTypeForVariant(v: TListPageAlertVariant): MessageBarType {
  if (v === 'success') return MessageBarType.success;
  if (v === 'warning') return MessageBarType.warning;
  if (v === 'error') return MessageBarType.error;
  return MessageBarType.info;
}

const DEFAULT_ICON: Record<TListPageAlertVariant, string> = {
  info: 'Info',
  success: 'CheckMark',
  warning: 'Warning',
  error: 'ErrorBadge',
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
        styles={{ root: { marginBottom: 8 } }}
      >
        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#605e5c' } }}>
          Alerta
        </Text>
        <ActionButton
          iconProps={{ iconName: 'Settings' }}
          onClick={onConfigure}
          styles={{ root: { height: 28, color: '#0078d4' } }}
        >
          Configurar
        </ActionButton>
      </Stack>
    ) : null;

  if (dismissed && c.dismissible) {
    return (
      <Stack tokens={{ childrenGap: 6 }}>
        {toolbar}
        {countLoading && <Spinner label="A avaliar regras de contagem…" />}
      </Stack>
    );
  }

  const emphasizedStyles = c.emphasized
    ? {
        root: {
          borderWidth: 1,
          borderStyle: 'solid',
          borderColor: '#8a8886',
          borderRadius: 4,
        },
      }
    : undefined;

  const inner = (
    <Stack tokens={{ childrenGap: 8 }}>
      {countErr && (
        <Text variant="small" styles={{ root: { color: '#a4262c' } }}>
          Contagem: {countErr}
        </Text>
      )}
      {c.title.trim() ? (
        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
          {c.title}
        </Text>
      ) : null}
      {c.message.trim() ? (
        <Text variant="small" styles={{ root: { lineHeight: 1.5 } }}>
          {c.message}
        </Text>
      ) : (
        onConfigure !== undefined && (
          <Text variant="small" styles={{ root: { color: '#a19f9d', fontStyle: 'italic' } }}>
            Sem mensagem — configure o bloco
          </Text>
        )
      )}
      {hasLink ? (
        <Link href={c.linkUrl.trim()} target="_blank" rel="noopener noreferrer">
          {c.linkText.trim()}
        </Link>
      ) : null}
    </Stack>
  );

  return (
    <>
      {toolbar}
      {countLoading && !dismissed ? (
        <Spinner label="A avaliar regras de contagem…" styles={{ root: { marginBottom: 8 } }} />
      ) : null}
      <MessageBar
        messageBarType={messageBarTypeForVariant(c.variant)}
        isMultiline
        onDismiss={c.dismissible ? () => setDismissed(true) : undefined}
        dismissButtonAriaLabel="Fechar"
        messageBarIconProps={iconOverride ? { iconName: iconOverride } : { iconName: DEFAULT_ICON[c.variant] }}
        styles={emphasizedStyles}
      >
        {inner}
      </MessageBar>
    </>
  );
};
