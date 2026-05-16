import * as React from 'react';
import { Label, Stack, Text, useTheme } from '@fluentui/react';
import {
  getFormControlBorderRadius,
  getFormFieldHelpTextRootStyle,
  getRequiredEmptyBorderColor,
} from '../../core/formManager/formControlFluentStyles';

export interface IMultilineReadonlyHtmlProps {
  label: string;
  required?: boolean;
  html: string;
  help?: React.ReactNode;
  showReqEmpty?: boolean;
  showLabel?: boolean;
}

export const MultilineReadonlyHtml: React.FC<IMultilineReadonlyHtmlProps> = ({
  label,
  required,
  html,
  help,
  showReqEmpty,
  showLabel = true,
}) => {
  const theme = useTheme();
  const r = getFormControlBorderRadius(theme);
  const reqBorder = getRequiredEmptyBorderColor(theme);
  const helpRoot = getFormFieldHelpTextRootStyle(theme);
  return (
    <Stack
      tokens={{ childrenGap: 6 }}
      styles={{
        root: {
          marginBottom: 12,
          ...(showReqEmpty
            ? {
                borderLeft: `3px solid ${reqBorder}`,
                paddingLeft: 8,
                paddingTop: 2,
                paddingBottom: 2,
              }
            : {}),
        },
      }}
    >
      {showLabel ? <Label required={required}>{label}</Label> : null}
      <div
        className="dinamic-sp-rich-note"
        title={!showLabel ? label : undefined}
        style={{
          padding: '8px 10px',
          border: `1px solid ${theme.palette.neutralLight}`,
          borderRadius: r,
          background: theme.palette.white,
          minHeight: 40,
          lineHeight: 1.5,
          cursor: 'not-allowed',
        }}
        dangerouslySetInnerHTML={{ __html: html }}
      />
      {help ? (
        typeof help === 'string' ? (
          <Text variant="small" styles={{ root: helpRoot }}>
            {help}
          </Text>
        ) : (
          help
        )
      ) : null}
    </Stack>
  );
};
