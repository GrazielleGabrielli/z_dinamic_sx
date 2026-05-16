import * as React from 'react';
import type { IDropdownOption, IDropdownStyles, ITheme } from '@fluentui/react';
import {
  FORM_FIELD_CURSOR_DISABLED,
  getFormControlBorderRadius,
  getRequiredEmptyBorderColor,
} from '../../core/formManager/formControlFluentStyles';

export function multiSelectDropdownStyles(
  theme: ITheme,
  showReq: boolean | undefined,
  disabled?: boolean
): Partial<IDropdownStyles> {
  const dropdown: Record<string, string | number> = {};
  const r = getFormControlBorderRadius(theme);
  if (showReq === true) {
    Object.assign(dropdown, {
      borderColor: getRequiredEmptyBorderColor(theme),
      borderWidth: 1,
      borderStyle: 'solid' as const,
      borderRadius: r,
    });
  }
  if (disabled) {
    const text = theme.palette.neutralPrimary;
    Object.assign(dropdown, { color: text, opacity: 1, cursor: FORM_FIELD_CURSOR_DISABLED });
  }
  const base: Partial<IDropdownStyles> = {
    title: {
      height: 'auto',
      minHeight: 32,
      lineHeight: '20px',
      whiteSpace: 'normal',
      overflow: 'visible',
      display: 'flex',
      flexWrap: 'wrap',
      alignItems: 'center',
      paddingTop: 4,
      paddingBottom: 4,
      paddingRight: 32,
      ...(disabled
        ? {
            color: theme.palette.neutralPrimary,
            opacity: 1,
            WebkitTextFillColor: theme.palette.neutralPrimary,
            cursor: FORM_FIELD_CURSOR_DISABLED,
          }
        : {}),
    },
    caretDownWrapper: {
      height: 'auto',
      minHeight: 32,
      alignSelf: 'stretch',
      display: 'flex',
      alignItems: 'center',
      top: 0,
      ...(disabled ? { cursor: FORM_FIELD_CURSOR_DISABLED } : {}),
    },
  };
  if (Object.keys(dropdown).length > 0) {
    base.dropdown = dropdown as IDropdownStyles['dropdown'];
  }
  if (disabled) {
    base.caretDown = {
      color: theme.palette.neutralSecondary,
      opacity: 1,
      cursor: FORM_FIELD_CURSOR_DISABLED,
    };
  }
  return base;
}

export function renderMultiSelectDropdownTitle(
  theme: ITheme,
  options?: IDropdownOption[] | null,
  disabled?: boolean
): React.ReactElement | null {
  if (!options || options.length === 0) {
    return null;
  }
  const bg = disabled
    ? (theme.palette.neutralLighterAlt ?? theme.palette.white)
    : (theme.palette.themeLighterAlt ?? theme.palette.themeLighter);
  const fg = disabled ? theme.palette.neutralPrimary : theme.palette.themePrimary;
  const border = disabled ? (theme.palette.neutralLight ?? '#edebe9') : theme.palette.themeLight;
  const r = getFormControlBorderRadius(theme);
  const fs = theme.fonts.small;
  return (
    <span
      style={{
        display: 'inline-flex',
        flexWrap: 'wrap',
        gap: 6,
        alignItems: 'center',
        maxWidth: '100%',
        ...(disabled ? { cursor: FORM_FIELD_CURSOR_DISABLED } : {}),
      }}
    >
      {options.map((o) => (
        <span
          key={String(o.key)}
          title={o.text}
          style={{
            padding: '2px 8px',
            borderRadius: r,
            background: bg,
            color: fg,
            border: `1px solid ${border}`,
            fontSize: fs.fontSize,
            lineHeight: fs.lineHeight,
            maxWidth: '100%',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
          }}
        >
          {o.text}
        </span>
      ))}
    </span>
  );
}
