import * as React from 'react';
import type { IDropdownOption, IDropdownStyles, ITheme } from '@fluentui/react';

const REQ_EMPTY_BORDER = '#a4262c';

export function multiSelectDropdownStyles(showReq: boolean | undefined): Partial<IDropdownStyles> {
  const reqBorder =
    showReq === true
      ? {
          dropdown: {
            borderColor: REQ_EMPTY_BORDER,
            borderWidth: 1,
            borderStyle: 'solid' as const,
            borderRadius: 2,
          },
        }
      : {};
  return {
    ...reqBorder,
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
    },
    caretDownWrapper: {
      height: 'auto',
      minHeight: 32,
      alignSelf: 'stretch',
      display: 'flex',
      alignItems: 'center',
      top: 0,
    },
  };
}

export function renderMultiSelectDropdownTitle(
  theme: ITheme,
  options?: IDropdownOption[] | null
): React.ReactElement | null {
  if (!options || options.length === 0) {
    return null;
  }
  const bg = theme.palette.themeLighterAlt ?? theme.palette.themeLighter;
  const fg = theme.palette.themePrimary;
  const border = theme.palette.themeLight;
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
