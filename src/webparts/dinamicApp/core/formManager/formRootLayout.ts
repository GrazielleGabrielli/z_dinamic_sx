import type { CSSProperties } from 'react';
import type { IFormManagerConfig } from '../config/types/formManager';

function formRootPaddingStyle(fm: IFormManagerConfig): Pick<CSSProperties, 'padding'> | undefined {
  const n = fm.formRootPaddingPx;
  if (typeof n !== 'number' || !isFinite(n) || n <= 0) return undefined;
  const px = Math.min(160, Math.max(1, Math.round(n)));
  return { padding: px };
}

export function resolveFormRootLayoutStyles(fm: IFormManagerConfig): {
  outer: CSSProperties;
  inner: CSSProperties;
} {
  const legacy =
    fm.formRootWidthMode === undefined &&
    fm.formRootWidthPercent === undefined &&
    fm.formRootHorizontalAlign === undefined;

  const align = fm.formRootHorizontalAlign ?? 'start';
  const justify: CSSProperties['justifyContent'] =
    align === 'end' ? 'flex-end' : align === 'center' ? 'center' : 'flex-start';
  const pad = formRootPaddingStyle(fm);

  if (legacy) {
    return {
      outer: { width: '100%', display: 'flex', justifyContent: 'flex-start', boxSizing: 'border-box' },
      inner: {
        width: '100%',
        maxWidth: 720,
        minWidth: 0,
        boxSizing: 'border-box',
        ...(pad ?? {}),
      },
    };
  }

  const mode = fm.formRootWidthMode ?? 'percent';
  const pct = Math.min(100, Math.max(1, fm.formRootWidthPercent ?? 100));

  if (mode === 'full') {
    return {
      outer: { width: '100%', display: 'flex', justifyContent: justify, boxSizing: 'border-box' },
      inner: {
        width: '100%',
        maxWidth: '100%',
        minWidth: 0,
        boxSizing: 'border-box',
        ...(pad ?? {}),
      },
    };
  }

  return {
    outer: { width: '100%', display: 'flex', justifyContent: justify, boxSizing: 'border-box' },
    inner: {
      width: `${pct}%`,
      maxWidth: '100%',
      minWidth: 0,
      boxSizing: 'border-box',
      ...(pad ?? {}),
    },
  };
}
