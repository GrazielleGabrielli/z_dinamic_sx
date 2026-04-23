import * as React from 'react';
import { Stack, getTheme } from '@fluentui/react';

export function parseMultiChoiceLabels(raw: unknown): string[] {
  if (raw == null || raw === '') return [];
  if (Array.isArray(raw)) {
    return raw.map((v) => String(v).trim()).filter((s) => s.length > 0);
  }
  if (typeof raw === 'string') {
    const s = raw.trim();
    if (!s) return [];
    if (s.charAt(0) === '[' && s.charAt(s.length - 1) === ']') {
      try {
        const parsed = JSON.parse(s) as unknown;
        if (Array.isArray(parsed)) {
          return parsed.map((v) => String(v).trim()).filter((t) => t.length > 0);
        }
      } catch {
        return [];
      }
    }
    if (s.indexOf(';') !== -1) {
      return s.split(';').map((t) => t.trim()).filter((t) => t.length > 0);
    }
    return [s];
  }
  return [String(raw)];
}

export function MultiChoiceBadges(props: { labels: string[]; emptyFallback: string }): React.ReactElement {
  const { labels, emptyFallback } = props;
  if (labels.length === 0) return <>{emptyFallback}</>;
  const theme = getTheme();
  const bg = theme.palette.themeLighter;
  const fg = theme.palette.themeDark;
  const border = theme.palette.themePrimary;
  const radius =
    theme.effects && typeof theme.effects.roundedCorner2 === 'string'
      ? theme.effects.roundedCorner2
      : '2px';
  return (
    <Stack horizontal wrap tokens={{ childrenGap: 6 }}>
      {labels.map((label, i) => (
        <span
          key={`${i}:${label}`}
          style={{
            display: 'inline-block',
            padding: '2px 8px',
            borderRadius: radius,
            background: bg,
            color: fg,
            border: `1px solid ${border}`,
            fontSize: theme.fonts.small?.size ?? 12,
            fontWeight: theme.fonts.small?.fontWeight as string | number,
            maxWidth: '100%',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
          }}
          title={label}
        >
          {label}
        </span>
      ))}
    </Stack>
  );
}
