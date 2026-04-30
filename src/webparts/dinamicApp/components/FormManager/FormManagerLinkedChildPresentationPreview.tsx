import * as React from 'react';
import { useMemo } from 'react';
import { Stack, Text } from '@fluentui/react';
import type { IFieldMetadata } from '../../../../services';
import type { IFormLinkedChildFormConfig, TLinkedChildRowsPresentationKind } from '../../core/config/types/formManager';
import { getLinkedChildOrderedFieldConfigs } from '../../core/formManager/formLinkedChildSync';

function previewLabels(cfg: IFormLinkedChildFormConfig, meta: IFieldMetadata[]): string[] {
  const ordered = getLinkedChildOrderedFieldConfigs(cfg);
  const byName = new Map(meta.map((m) => [m.InternalName, m]));
  const out: string[] = [];
  for (let i = 0; i < ordered.length && i < 3; i++) {
    const fc = ordered[i];
    const m = byName.get(fc.internalName);
    out.push((fc.label ?? m?.Title ?? fc.internalName).trim() || fc.internalName);
  }
  while (out.length < 2) {
    out.push(`Campo ${out.length + 1}`);
  }
  return out.slice(0, 3);
}

function MockFieldLine(props: { label: string; dense?: boolean }): JSX.Element {
  const { label, dense } = props;
  return (
    <Stack tokens={{ childrenGap: dense ? 2 : 4 }}>
      <Text variant="small" styles={{ root: { color: '#605e5c', fontWeight: 600 } }}>
        {label}
      </Text>
      <div
        style={{
          height: dense ? 22 : 28,
          borderRadius: 2,
          border: '1px solid #c8c6c4',
          background: '#ffffff',
        }}
      />
    </Stack>
  );
}

function PreviewBlock(props: {
  title: string;
  kind: TLinkedChildRowsPresentationKind;
  labels: string[];
}): JSX.Element {
  const { title, kind, labels } = props;
  const [a, b, c] = [labels[0] ?? 'A', labels[1] ?? 'B', labels[2]];

  if (kind === 'table') {
    return (
      <Stack tokens={{ childrenGap: 6 }}>
        <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
          {title}
        </Text>
        <div style={{ overflowX: 'auto', border: '1px solid #edebe9', borderRadius: 2 }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
            <thead>
              <tr style={{ background: '#f3f2f1' }}>
                <th style={{ textAlign: 'left', padding: '6px 8px', borderBottom: '2px solid #edebe9', width: 36 }}>
                  {' '}
                </th>
                {labels.map((h, i) => (
                  <th
                    key={i}
                    style={{
                      textAlign: 'left',
                      padding: '6px 8px',
                      borderBottom: '2px solid #edebe9',
                      borderLeft: '1px solid #edebe9',
                      fontWeight: 600,
                    }}
                  >
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              <tr>
                <td
                  style={{
                    padding: '6px 8px',
                    borderBottom: '1px solid #edebe9',
                    background: '#faf9f8',
                    whiteSpace: 'nowrap',
                  }}
                >
                  <Text variant="tiny" styles={{ root: { fontWeight: 600 } }}>
                    1
                  </Text>
                </td>
                {labels.map((_, i) => (
                  <td
                    key={i}
                    style={{
                      padding: '6px 8px',
                      borderBottom: '1px solid #edebe9',
                      borderLeft: '1px solid #edebe9',
                      verticalAlign: 'top',
                    }}
                  >
                    <div style={{ height: 22, border: '1px solid #c8c6c4', borderRadius: 2, background: '#fff' }} />
                  </td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>
      </Stack>
    );
  }

  const surface: React.CSSProperties =
    kind === 'cards'
      ? {
          border: '1px solid #e1dfdd',
          borderRadius: 8,
          padding: 12,
          background: '#ffffff',
          boxShadow: '0 1.6px 3.6px rgba(0, 0, 0, 0.09)',
        }
      : kind === 'compact'
        ? {
            border: '1px solid #edebe9',
            borderRadius: 4,
            padding: 8,
            background: '#faf9f8',
          }
        : {
            border: '1px solid #edebe9',
            borderRadius: 4,
            padding: 12,
            background: '#faf9f8',
          };

  return (
    <Stack tokens={{ childrenGap: 6 }}>
      <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
        {title}
      </Text>
      <div style={surface}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <Text variant="tiny" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
            Linha 1
          </Text>
        </Stack>
        <Stack tokens={{ childrenGap: kind === 'compact' ? 6 : 8 }} styles={{ root: { marginTop: 8 } }}>
          <MockFieldLine label={a} dense={kind === 'compact'} />
          <MockFieldLine label={b} dense={kind === 'compact'} />
          {c && labels.length > 2 ? <MockFieldLine label={c} dense={kind === 'compact'} /> : null}
        </Stack>
      </div>
    </Stack>
  );
}

const PREVIEW_TITLE_BY_KIND: Record<TLinkedChildRowsPresentationKind, string> = {
  stack: 'Blocos (em coluna)',
  table: 'Tabela',
  compact: 'Compacto',
  cards: 'Cartões',
};

export interface IFormManagerLinkedChildPresentationPreviewProps {
  cfg: IFormLinkedChildFormConfig;
  fieldMeta: IFieldMetadata[];
  /** Modo escolhido no dropdown «Apresentação das linhas». */
  presentationKind: TLinkedChildRowsPresentationKind;
}

export const FormManagerLinkedChildPresentationPreview: React.FC<
  IFormManagerLinkedChildPresentationPreviewProps
> = ({ cfg, fieldMeta, presentationKind }) => {
  const labels = useMemo(() => previewLabels(cfg, fieldMeta), [cfg, fieldMeta]);
  const title = PREVIEW_TITLE_BY_KIND[presentationKind] ?? PREVIEW_TITLE_BY_KIND.stack;

  return (
    <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 8 } }}>
  
      <PreviewBlock title={title} kind={presentationKind} labels={labels} />
    </Stack>
  );
};
