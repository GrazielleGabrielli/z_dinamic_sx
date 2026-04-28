import * as React from 'react';
import { ActionButton, DefaultButton, PrimaryButton, Stack, Text } from '@fluentui/react';
import type { IListPageButtonItemConfig, IListPageButtonsBlockConfig } from '../../core/config/types';
import { defaultButtonsConfig } from '../../core/listPage/listPageBlockConfigUtils';

export interface IListPageButtonsBlockProps {
  buttons?: IListPageButtonsBlockConfig;
  onConfigure?: () => void;
}

function navigateItem(it: IListPageButtonItemConfig): void {
  if (it.actionKind === 'reload') {
    window.location.reload();
    return;
  }
  const u = (it.url ?? '').trim();
  if (!u) return;
  if (it.openInNewTab === true) {
    window.open(u, '_blank', 'noopener,noreferrer');
  } else {
    window.location.assign(u);
  }
}

function parseCssToStyle(css: string | undefined): React.CSSProperties | undefined {
  if (!css?.trim()) return undefined;
  const style: Record<string, string> = {};
  css.split(';').forEach((decl) => {
    const idx = decl.indexOf(':');
    if (idx < 0) return;
    const prop = decl.slice(0, idx).trim();
    const val = decl.slice(idx + 1).trim();
    if (!prop || !val) return;
    const camel = prop.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());
    style[camel] = val;
  });
  return Object.keys(style).length > 0 ? (style as React.CSSProperties) : undefined;
}

const ALIGN_MAP: Record<string, 'start' | 'center' | 'end'> = {
  left: 'start',
  center: 'center',
  right: 'end',
};

export const ListPageButtonsBlock: React.FC<IListPageButtonsBlockProps> = ({
  buttons: raw,
  onConfigure,
}) => {
  const cfg = raw ?? defaultButtonsConfig();
  const items = cfg.items ?? [];
  const gap = cfg.gap ?? 8;
  const align = ALIGN_MAP[cfg.align ?? 'left'] ?? 'start';
  const containerStyle = parseCssToStyle(cfg.containerCss);

  const toolbar =
    onConfigure !== undefined ? (
      <Stack
        horizontal
        horizontalAlign="space-between"
        verticalAlign="center"
        styles={{ root: { marginBottom: 8 } }}
      >
        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#605e5c' } }}>
          Botões
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

  if (items.length === 0 && onConfigure === undefined) {
    return null;
  }

  return (
    <>
      {toolbar}
      <div className="dinamicSxButtons" style={{ display: 'flex', flexWrap: 'wrap', gap, justifyContent: align, alignItems: 'center', ...containerStyle }}>
        {items.map((it) => {
          const btnStyle = parseCssToStyle(it.css);
          const iconProps = it.iconName ? { iconName: it.iconName } : undefined;
          const btn =
            it.variant === 'primary' ? (
              <PrimaryButton
                key={it.id}
                text={it.label}
                iconProps={iconProps}
                onClick={() => navigateItem(it)}
                styles={{ root: { height: 32 } }}
              />
            ) : (
              <DefaultButton
                key={it.id}
                text={it.label}
                iconProps={iconProps}
                onClick={() => navigateItem(it)}
                styles={{ root: { height: 32 } }}
              />
            );
          return btnStyle ? (
            <span key={it.id} style={btnStyle}>
              {btn}
            </span>
          ) : (
            <React.Fragment key={it.id}>{btn}</React.Fragment>
          );
        })}
      </div>
    </>
  );
};
