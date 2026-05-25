import type {
  IDropdownStyles,
  IStyle,
  IStyleFunctionOrObject,
  ITextFieldStyleProps,
  ITextFieldStyles,
  ITheme,
} from '@fluentui/react';

export const FORM_FIELD_CURSOR_DISABLED = 'not-allowed';

export type TFluentTextFieldStyles = IStyleFunctionOrObject<ITextFieldStyleProps, ITextFieldStyles>;

export function asFluentTextFieldStyles(
  styles: Partial<ITextFieldStyles> | undefined
): TFluentTextFieldStyles | undefined {
  return styles as TFluentTextFieldStyles | undefined;
}

export function getFormControlBorderRadius(theme: ITheme): number {
  const r = theme.effects?.roundedCorner2;
  if (typeof r === 'number' && isFinite(r)) return r;
  if (typeof r === 'string') {
    const n = parseInt(r, 10);
    if (!isNaN(n)) return n;
  }
  return 2;
}

export function getRequiredEmptyBorderColor(theme: ITheme): string {
  return theme.semanticColors.errorText;
}

export function mergeFormTextFieldStyles(
  ...parts: Array<Partial<ITextFieldStyles> | undefined>
): TFluentTextFieldStyles | undefined {
  const out: Partial<ITextFieldStyles> = {};
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    if (!p) continue;
    const keys = Object.keys(p) as (keyof ITextFieldStyles)[];
    for (let k = 0; k < keys.length; k++) {
      const key = keys[k];
      const v = p[key];
      if (v === undefined) continue;
      const prev = out[key];
      if (
        prev &&
        typeof prev === 'object' &&
        !Array.isArray(prev) &&
        v &&
        typeof v === 'object' &&
        !Array.isArray(v)
      ) {
        (out as Record<string, unknown>)[key as string] = {
          ...(prev as object),
          ...(v as object),
        };
      } else {
        (out as Record<string, unknown>)[key as string] = v;
      }
    }
  }
  return asFluentTextFieldStyles(Object.keys(out).length ? out : undefined);
}

export function getFormFieldDescriptionSlotStyles(theme: ITheme): Partial<ITextFieldStyles> {
  const small = theme.fonts.small ?? theme.fonts.medium;
  return {
    description: {
      color: theme.palette.neutralSecondary,
      fontSize: small.fontSize,
      lineHeight: small.lineHeight,
    },
  };
}

export function getFormFieldHelpTextRootStyle(theme: ITheme): IStyle {
  const small = theme.fonts.small ?? theme.fonts.medium;
  return {
    color: theme.palette.neutralSecondary,
    fontSize: small.fontSize,
    lineHeight: small.lineHeight,
  };
}

export function getFormTextFieldStyles(
  theme: ITheme,
  opts: { requiredEmpty: boolean; disabled?: boolean }
): Partial<ITextFieldStyles> | undefined {
  const r = getFormControlBorderRadius(theme);
  const reqBorder = getRequiredEmptyBorderColor(theme);
  const fieldGroupMerge: Record<string, unknown> = {};
  if (opts.requiredEmpty) {
    Object.assign(fieldGroupMerge, {
      borderColor: reqBorder,
      borderWidth: 1,
      borderStyle: 'solid',
      borderRadius: r,
    });
  }
  if (opts.disabled) {
    Object.assign(fieldGroupMerge, { cursor: FORM_FIELD_CURSOR_DISABLED });
  }
  const out: Partial<ITextFieldStyles> = {};
  if (Object.keys(fieldGroupMerge).length) {
    out.fieldGroup = fieldGroupMerge as IStyle;
  }
  if (opts.disabled) {
    const text = theme.palette.neutralPrimary;
    const ph = theme.palette.neutralSecondary;
    out.root = { cursor: FORM_FIELD_CURSOR_DISABLED };
    out.icon = { cursor: FORM_FIELD_CURSOR_DISABLED };
    out.field = {
      color: text,
      WebkitTextFillColor: text,
      opacity: 1,
      cursor: FORM_FIELD_CURSOR_DISABLED,
      selectors: {
        '::placeholder': {
          color: ph,
          opacity: 1,
        },
      },
    };
  }
  return Object.keys(out).length ? out : undefined;
}

export function getFormDropdownStyles(
  theme: ITheme,
  opts: { requiredEmpty: boolean | undefined; disabled?: boolean }
): Partial<IDropdownStyles> | undefined {
  const r = getFormControlBorderRadius(theme);
  const reqBorder = getRequiredEmptyBorderColor(theme);
  const out: Partial<IDropdownStyles> = {};
  const dropdown: Record<string, unknown> = {};
  if (opts.requiredEmpty) {
    Object.assign(dropdown, {
      borderColor: reqBorder,
      borderWidth: 1,
      borderStyle: 'solid',
      borderRadius: r,
    });
  }
  if (opts.disabled) {
    const text = theme.palette.neutralPrimary;
    const caret = theme.palette.neutralSecondary;
    out.title = {
      color: text,
      opacity: 1,
      WebkitTextFillColor: text,
      cursor: FORM_FIELD_CURSOR_DISABLED,
    };
    out.caretDownWrapper = { cursor: FORM_FIELD_CURSOR_DISABLED };
    out.caretDown = { color: caret, opacity: 1, cursor: FORM_FIELD_CURSOR_DISABLED };
    Object.assign(dropdown, { color: text, opacity: 1, cursor: FORM_FIELD_CURSOR_DISABLED });
  }
  if (Object.keys(dropdown).length) {
    out.dropdown = dropdown as IDropdownStyles['dropdown'];
  }
  return Object.keys(out).length ? out : undefined;
}

export function getFormConfirmPromptTextFieldStyles(
  theme: ITheme,
  opts: { disabled?: boolean; modalSurface?: boolean }
): TFluentTextFieldStyles {
  const dis = opts.disabled === true;
  const r = getFormControlBorderRadius(theme);
  const fieldGroup: IStyle = {
    borderRadius: r,
    border: `1px solid ${theme.palette.neutralQuaternaryAlt}`,
    backgroundColor: theme.palette.white,
    ':hover': { borderColor: theme.palette.neutralTertiaryAlt },
    selectors: {
      '&.ms-TextField-fieldGroup': { borderRadius: r },
    },
    ...(dis ? { cursor: FORM_FIELD_CURSOR_DISABLED } : {}),
  };
  return {
    root: {
      marginBottom: 0,
      ...(opts.modalSurface ? { width: '100%' } : {}),
      ...(dis ? { cursor: FORM_FIELD_CURSOR_DISABLED } : {}),
    },
    label: {
      root: {
        fontWeight: '600',
        color: theme.palette.neutralPrimary,
        marginBottom: opts.modalSurface ? 8 : 6,
        ...(opts.modalSurface ? { fontSize: 14, lineHeight: '20px' } : {}),
      },
    },
    fieldGroup,
    field: { borderRadius: r, ...(dis ? { cursor: FORM_FIELD_CURSOR_DISABLED } : {}) },
    ...(dis ? { icon: { cursor: FORM_FIELD_CURSOR_DISABLED } } : {}),
    ...(opts.modalSurface ? { wrapper: { width: '100%' } } : {}),
  } as TFluentTextFieldStyles;
}

export function getFormConfirmPromptDropdownStyles(
  theme: ITheme,
  opts: { disabled?: boolean }
): Partial<IDropdownStyles> {
  const dis = opts.disabled === true;
  const r = getFormControlBorderRadius(theme);
  return {
    dropdown: {
      borderRadius: r,
      border: `1px solid ${theme.palette.neutralQuaternaryAlt}`,
      ...(dis ? { cursor: FORM_FIELD_CURSOR_DISABLED } : {}),
    },
    ...(dis
      ? {
          title: { cursor: FORM_FIELD_CURSOR_DISABLED },
          caretDownWrapper: { cursor: FORM_FIELD_CURSOR_DISABLED },
          caretDown: { cursor: FORM_FIELD_CURSOR_DISABLED },
        }
      : {}),
  };
}
