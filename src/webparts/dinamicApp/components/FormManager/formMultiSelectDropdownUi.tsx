import * as React from 'react';
import type { IDropdownOption, IDropdownStyles, ITheme } from '@fluentui/react';

const REQ_EMPTY_BORDER = '#a4262c';

const FORM_FIELD_CURSOR_DISABLED = 'not-allowed';

export function multiSelectDropdownStyles(
  showReq: boolean | undefined,
  disabled?: boolean
): Partial<IDropdownStyles> {
  const dropdown: Record<string, string | number> = {};
  if (showReq === true) {
    Object.assign(dropdown, {
      borderColor: REQ_EMPTY_BORDER,
      borderWidth: 1,
      borderStyle: 'solid' as const,
      borderRadius: 2,
    });
  }
  if (disabled) {
    Object.assign(dropdown, { color: '#201f1e', opacity: 1, cursor: FORM_FIELD_CURSOR_DISABLED });
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
            color: '#201f1e',
            opacity: 1,
            WebkitTextFillColor: '#201f1e',
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
    base.caretDown = { color: '#605e5c', opacity: 1, cursor: FORM_FIELD_CURSOR_DISABLED };
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
  const fg = disabled ? '#201f1e' : theme.palette.themePrimary;
  const border = disabled ? (theme.palette.neutralLight ?? '#edebe9') : theme.palette.themeLight;
  const r = theme.effects?.roundedCorner2 ?? 2;
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
