import * as React from 'react';
import { PrimaryButton, Stack, Text, ActionButton } from '@fluentui/react';
import type { IListPageBannerBlockConfig } from '../../core/config/types';
import { defaultBannerConfig } from '../../core/listPage/listPageBlockConfigUtils';

export interface IListPageBannerBlockProps {
  banner?: IListPageBannerBlockConfig;
  onConfigure?: () => void;
}

export const ListPageBannerBlock: React.FC<IListPageBannerBlockProps> = ({ banner: raw, onConfigure }) => {
  const c = raw ?? defaultBannerConfig();
  const h = Math.max(80, Math.min(800, c.heightPx));
  const overlay = Math.max(0, Math.min(1, c.overlayOpacity));
  const align = c.contentAlign === 'left' || c.contentAlign === 'right' ? c.contentAlign : 'center';
  const textAlign = align === 'left' ? 'left' : align === 'right' ? 'right' : 'center';
  const hasImg = Boolean(c.imageUrl.trim());

  const toolbar =
    onConfigure !== undefined ? (
      <Stack
        horizontal
        horizontalAlign="space-between"
        verticalAlign="center"
        styles={{ root: { marginBottom: 8 } }}
      >
        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#605e5c' } }}>
          Banner
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

  const inner = (
    <div
      style={{
        position: 'relative',
        zIndex: 1,
        padding: '24px 28px',
        textAlign,
        maxWidth: 720,
        margin: align === 'center' ? '0 auto' : align === 'right' ? '0 0 0 auto' : '0',
      }}
    >
      {c.title.trim() ? (
        <div style={{ fontSize: 28, fontWeight: 700, color: '#fff', lineHeight: 1.2, marginBottom: 8 }}>
          {c.title}
        </div>
      ) : null}
      {c.subtitle.trim() ? (
        <div style={{ fontSize: 15, color: 'rgba(255,255,255,0.92)', lineHeight: 1.45, marginBottom: 16 }}>
          {c.subtitle}
        </div>
      ) : null}
      {c.showButton && c.buttonText.trim() && c.linkUrl.trim() ? (
        c.openInNewTab ? (
          <a
            href={c.linkUrl.trim()}
            target="_blank"
            rel="noopener noreferrer"
            style={{ textDecoration: 'none', display: 'inline-block', marginTop: 4 }}
          >
            <PrimaryButton text={c.buttonText.trim()} styles={{ root: { marginTop: 0 } }} />
          </a>
        ) : (
          <PrimaryButton
            text={c.buttonText.trim()}
            href={c.linkUrl.trim()}
            styles={{ root: { marginTop: 4 } }}
          />
        )
      ) : null}
    </div>
  );

  const shell = (
    <div
      className="dinamicSxBanner"
      style={{
        position: 'relative',
        minHeight: h,
        borderRadius: 4,
        overflow: 'hidden',
        display: 'flex',
        flexDirection: 'column',
        justifyContent: 'center',
        background: hasImg ? '#323130' : 'linear-gradient(135deg, #0078d4 0%, #004578 100%)',
      }}
      role="region"
      aria-label={c.imageAlt.trim() || c.title.trim() || 'Banner'}
    >
      {hasImg ? (
        <img
          src={c.imageUrl.trim()}
          alt={c.imageAlt}
          style={{
            position: 'absolute',
            inset: 0,
            width: '100%',
            height: '100%',
            objectFit: 'cover',
          }}
        />
      ) : null}
      <div
        style={{
          position: 'absolute',
          inset: 0,
          background: `rgba(0,0,0,${overlay})`,
          pointerEvents: 'none',
        }}
      />
      {inner}
    </div>
  );

  if (!c.showButton && c.linkUrl.trim()) {
    return (
      <>
        {toolbar}
        <a
          href={c.linkUrl.trim()}
          target={c.openInNewTab ? '_blank' : undefined}
          rel={c.openInNewTab ? 'noreferrer noopener' : undefined}
          style={{ textDecoration: 'none', color: 'inherit', display: 'block' }}
        >
          {shell}
        </a>
      </>
    );
  }

  return (
    <>
      {toolbar}
      {shell}
    </>
  );
};
