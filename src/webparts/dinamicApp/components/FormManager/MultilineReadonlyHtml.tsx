import * as React from 'react';
import { Label, Stack, Text } from '@fluentui/react';

const REQ_EMPTY_BORDER = '#a4262c';

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
}) => (
  <Stack
    tokens={{ childrenGap: 6 }}
    styles={{
      root: {
        marginBottom: 12,
        ...(showReqEmpty
          ? {
              borderLeft: `3px solid ${REQ_EMPTY_BORDER}`,
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
        border: '1px solid #edebe9',
        borderRadius: 2,
        background: '#ffffff',
        minHeight: 40,
        lineHeight: 1.5,
        cursor: 'not-allowed',
      }}
      dangerouslySetInnerHTML={{ __html: html }}
    />
    {help ? (
      typeof help === 'string' ? (
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          {help}
        </Text>
      ) : (
        help
      )
    ) : null}
  </Stack>
);
