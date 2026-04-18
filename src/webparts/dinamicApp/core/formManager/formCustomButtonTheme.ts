import type { IButtonStyles, ITheme } from '@fluentui/react';
import type { IFormCustomButtonConfig, TFormCustomButtonPaletteSlot } from '../config/types/formManager';

export const FORM_CUSTOM_BUTTON_SLOT_LABELS: Record<TFormCustomButtonPaletteSlot, string> = {
  outline: 'Contorno (neutro)',
  themePrimary: 'Primária do tema',
  themeSecondary: 'Secundária do tema',
  themeTertiary: 'Terciária do tema',
  themeDark: 'Tema escuro',
  themeDarkAlt: 'Tema escuro (alt.)',
  themeDarker: 'Tema mais escuro',
  themeLight: 'Tema claro',
  themeLighter: 'Tema mais claro',
  themeLighterAlt: 'Tema muito claro',
};

const PALETTE_KEYS: readonly Exclude<TFormCustomButtonPaletteSlot, 'outline'>[] = [
  'themePrimary',
  'themeSecondary',
  'themeTertiary',
  'themeDark',
  'themeDarkAlt',
  'themeDarker',
  'themeLight',
  'themeLighter',
  'themeLighterAlt',
];

export const STEP_UI_FALLBACK_ACCENT_HEX = '#0078d4';

function parseHex(hex: string): { r: number; g: number; b: number } | undefined {
  const h = hex.replace('#', '').trim();
  if (h.length !== 3 && h.length !== 6) return undefined;
  const full = h.length === 3 ? h.split('').map((c) => c + c).join('') : h;
  if (full.length !== 6) return undefined;
  const n = parseInt(full, 16);
  if (Number.isNaN(n)) return undefined;
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}

function relativeLuminance(hex: string): number {
  const rgb = parseHex(hex);
  if (!rgb) return 0.5;
  const linear = (c: number): number => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  };
  const r = linear(rgb.r);
  const g = linear(rgb.g);
  const b = linear(rgb.b);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

function buttonLabelOnBackground(bg: string): string {
  return relativeLuminance(bg) > 0.55 ? '#323130' : '#ffffff';
}

function darkenHex(hex: string, factor: number): string {
  const rgb = parseHex(hex);
  if (!rgb) return hex;
  const d = (x: number): number => Math.max(0, Math.min(255, Math.round(x * factor)));
  const r = d(rgb.r);
  const g = d(rgb.g);
  const b = d(rgb.b);
  return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
}

/** Mistura canal a branco — `amount` 0–1 (ex.: 0,12 ≈ um pouco mais claro). */
function lightenHex(hex: string, amount: number): string {
  const rgb = parseHex(hex);
  if (!rgb) return hex;
  const a = Math.max(0, Math.min(1, amount));
  const mix = (c: number): number => Math.round(c + (255 - c) * a);
  const r = mix(rgb.r);
  const g = mix(rgb.g);
  const b = mix(rgb.b);
  return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
}

export function resolveFormCustomButtonPaletteSlot(btn: IFormCustomButtonConfig): TFormCustomButtonPaletteSlot {
  if (btn.themePaletteSlot) return btn.themePaletteSlot;
  return btn.appearance === 'primary' ? 'themePrimary' : 'outline';
}

export function paletteBgFromSlot(theme: ITheme, slot: Exclude<TFormCustomButtonPaletteSlot, 'outline'>): string {
  const p = theme.palette;
  const bySlot: Record<typeof slot, string | undefined> = {
    themePrimary: p.themePrimary,
    themeSecondary: p.themeSecondary,
    themeTertiary: p.themeTertiary,
    themeDark: p.themeDark,
    themeDarkAlt: p.themeDarkAlt,
    themeDarker: p.themeDarker,
    themeLight: p.themeLight,
    themeLighter: p.themeLighter,
    themeLighterAlt: p.themeLighterAlt,
  };
  const v = bySlot[slot];
  return typeof v === 'string' && v.length > 0 ? v : p.themePrimary;
}

export function hexToRgbaString(hex: string, alpha: number): string {
  const rgb = parseHex(hex);
  if (!rgb) return `rgba(0, 120, 212, ${alpha})`;
  return `rgba(${rgb.r},${rgb.g},${rgb.b},${alpha})`;
}

export function resolveStepUiAccentColor(
  theme: ITheme,
  slot: TFormCustomButtonPaletteSlot | undefined
): string {
  if (slot === undefined) {
    return paletteBgFromSlot(theme, 'themePrimary');
  }
  if (slot === 'outline') {
    return theme.palette.neutralSecondary;
  }
  return paletteBgFromSlot(theme, slot);
}

/** Cor de barra lateral do HTML gravado no registo de log (SharePoint). */
export function resolveActionLogPaletteAccentHex(
  theme: ITheme,
  slot: TFormCustomButtonPaletteSlot | undefined
): string {
  return resolveStepUiAccentColor(theme, slot ?? 'themePrimary');
}

export function getFilledPaletteButtonStyles(
  theme: ITheme,
  slot: Exclude<TFormCustomButtonPaletteSlot, 'outline'>
): IButtonStyles {
  const bg = paletteBgFromSlot(theme, slot);
  const fg = buttonLabelOnBackground(bg);
  const hoverBg = lightenHex(bg, 0.12);
  const hoverFg = buttonLabelOnBackground(hoverBg);
  const pressedBg = darkenHex(bg, 0.94);
  const pressedFg = buttonLabelOnBackground(pressedBg);
  const border = theme.palette.neutralSecondary;
  return {
    root: {
      backgroundColor: bg,
      borderColor: bg,
      color: fg,
      borderWidth: 1,
      selectors: {
        ':hover': {
          backgroundColor: hoverBg,
          borderColor: hoverBg,
          color: hoverFg,
        },
        ':active': {
          backgroundColor: pressedBg,
          borderColor: pressedBg,
          color: pressedFg,
        },
      },
    },
    rootDisabled: {
      backgroundColor: theme.palette.neutralLighter,
      borderColor: theme.palette.neutralLight,
      color: theme.palette.neutralSecondary,
      selectors: {
        ':hover': {},
        ':active': {},
      },
    },
    flexContainer: { height: '100%' },
    icon: { color: fg },
    label: { color: fg },
    splitButtonDivider: { backgroundColor: border },
  };
}

export const FORM_CUSTOM_BUTTON_THEME_SLOTS: readonly Exclude<TFormCustomButtonPaletteSlot, 'outline'>[] =
  PALETTE_KEYS;
