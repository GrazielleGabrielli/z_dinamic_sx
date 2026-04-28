import * as React from 'react';
import { Icon, Separator, Stack, Text, ActionButton } from '@fluentui/react';
import type { IListPageSectionTitleBlockConfig } from '../../core/config/types';
import { defaultSectionTitleConfig } from '../../core/listPage/listPageBlockConfigUtils';

export interface IListPageSectionTitleBlockProps {
  sectionTitle?: IListPageSectionTitleBlockConfig;
  onConfigure?: () => void;
}

const TITLE_PX: Record<string, number> = { sm: 18, md: 22, lg: 28 };
const SUBTITLE_PX: Record<string, number> = { sm: 13, md: 14, lg: 15 };

export const ListPageSectionTitleBlock: React.FC<IListPageSectionTitleBlockProps> = ({
  sectionTitle: raw,
  onConfigure,
}) => {
  const c = raw ?? defaultSectionTitleConfig();
  const align = c.align === 'center' || c.align === 'right' ? c.align : 'left';
  const textAlign = align === 'left' ? 'left' : align === 'right' ? 'right' : 'center';
  const flexDir = align === 'right' ? 'row-reverse' : 'row';
  const titlePx = TITLE_PX[c.size] ?? TITLE_PX.md;
  const subPx = SUBTITLE_PX[c.size] ?? SUBTITLE_PX.md;
  const mt = Math.max(0, Math.min(120, c.marginTopPx));
  const mb = Math.max(0, Math.min(120, c.marginBottomPx));
  const iconName = c.iconName.trim();
  const hasText = Boolean(c.title.trim() || c.subtitle.trim());
  if (!hasText && !iconName && onConfigure === undefined) {
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
          Título de seção
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

  return (
    <>
      {toolbar}
      <div className="dinamicSxSectionTitle" style={{ marginTop: mt, marginBottom: mb }}>
        <div
          style={{
            display: 'flex',
            flexDirection: flexDir,
            alignItems: align === 'center' ? 'center' : 'flex-start',
            justifyContent: align === 'center' ? 'center' : 'flex-start',
            gap: iconName ? 12 : 0,
            textAlign,
          }}
        >
          {iconName ? (
            <Icon
              iconName={iconName}
              styles={{
                root: {
                  fontSize: Math.round(titlePx * 1.15),
                  color: '#0078d4',
                  flexShrink: 0,
                  marginTop: align === 'center' ? 2 : 4,
                },
              }}
            />
          ) : null}
          <div style={{ flex: '1 1 auto', minWidth: 0 }}>
            {c.title.trim() ? (
              <div
                style={{
                  fontSize: titlePx,
                  fontWeight: 700,
                  color: '#323130',
                  lineHeight: 1.25,
                  marginBottom: c.subtitle.trim() ? 6 : 0,
                }}
              >
                {c.title}
              </div>
            ) : null}
            {c.subtitle.trim() ? (
              <div style={{ fontSize: subPx, color: '#605e5c', lineHeight: 1.45 }}>{c.subtitle}</div>
            ) : null}
            {!hasText && onConfigure !== undefined ? (
              <Text variant="small" styles={{ root: { color: '#a19f9d', fontStyle: 'italic' } }}>
                Sem título — configure o bloco
              </Text>
            ) : null}
          </div>
        </div>
        {c.showDivider ? (
          <Separator styles={{ root: { marginTop: hasText || iconName ? 14 : 8 } }} />
        ) : null}
      </div>
    </>
  );
};
