import * as React from 'react';
import { ActionButton, IconButton, Modal, Stack, Text } from '@fluentui/react';
import type { IDashboardConfig, IDynamicViewConfig, IListPageBlock, IListPageSection } from '../core/config/types';
import {
  effectiveConfigForListPageBlock,
  getEffectiveListPageSections,
  LEGACY_LIST_PAGE_LIST_BLOCK_ID,
  resolveDashboardForListBlock,
} from '../core/listPage/listPageLayoutUtils';

export type TPageEditableComponentPick =
  | { kind: 'listLayout' }
  | { kind: 'formManager' }
  | { kind: 'projectTable' }
  | { kind: 'listTable'; blockId: string }
  | { kind: 'dashboardCards'; blockId: string }
  | { kind: 'dashboardSeries'; blockId: string }
  | { kind: 'dashboardToCharts'; blockId: string }
  | { kind: 'contentBlock'; blockId: string };

export interface IPageEditableComponentsModalProps {
  isOpen: boolean;
  onDismiss: () => void;
  config: IDynamicViewConfig;
  onPick: (pick: TPageEditableComponentPick) => void;
}

function flattenBlocks(sections: IListPageSection[] | undefined): IListPageBlock[] {
  if (!sections || !Array.isArray(sections)) return [];
  const out: IListPageBlock[] = [];
  for (let si = 0; si < sections.length; si++) {
    const sec = sections[si];
    if (!sec?.columns || !Array.isArray(sec.columns)) continue;
    for (let ci = 0; ci < sec.columns.length; ci++) {
      const col = sec.columns[ci];
      if (!Array.isArray(col)) continue;
      for (let bi = 0; bi < col.length; bi++) {
        const b = col[bi];
        if (b) out.push(b);
      }
    }
  }
  return out;
}

function dashboardVisible(d: IDashboardConfig): boolean {
  return d.enabled && (d.dashboardType === 'charts' || (d.cardsCount ?? 0) > 0);
}

function listOrDashTitle(config: IDynamicViewConfig, block: IListPageBlock): string {
  return effectiveConfigForListPageBlock(config, block).dataSource.title?.trim() || '';
}

function blockContextLabel(config: IDynamicViewConfig, block: IListPageBlock): string {
  switch (block.type) {
    case 'dashboard':
      return listOrDashTitle(config, block) || 'Dashboard';
    case 'list': {
      const t = listOrDashTitle(config, block);
      return t || 'Lista';
    }
    case 'banner': {
      const t = block.banner?.title?.trim();
      return t ? `Banner «${t}»` : 'Banner';
    }
    case 'editor': {
      const t = block.editor?.title?.trim();
      return t ? `Texto «${t}»` : 'Bloco de texto';
    }
    case 'sectionTitle': {
      const t = block.sectionTitle?.title?.trim();
      return t ? `Título «${t}»` : 'Título de secção';
    }
    case 'alert': {
      const t = block.alert?.title?.trim();
      return t ? `Alerta «${t}»` : 'Alerta';
    }
    case 'buttons': {
      const first = block.buttons?.items?.[0]?.label?.trim();
      return first ? `Botões «${first}»` : 'Botões';
    }
    default:
      return block.type;
  }
}

export const PageEditableComponentsModal: React.FC<IPageEditableComponentsModalProps> = ({
  isOpen,
  onDismiss,
  config,
  onPick,
}) => {
  const rootDash = config.dashboard;

  const rows = React.useMemo(() => {
    const out: React.ReactNode[] = [];

    const pushRow = (key: string, line: string, pick: TPageEditableComponentPick): void => {
      out.push(
        <ActionButton
          key={key}
          iconProps={{ iconName: 'Settings' }}
          onClick={() => onPick(pick)}
          styles={{
            root: {
              width: '100%',
              justifyContent: 'flex-start',
              height: 'auto',
              minHeight: 44,
              padding: '10px 12px',
              borderBottom: '1px solid #edebe9',
            },
            flexContainer: { flexGrow: 1, textAlign: 'left' },
            label: { whiteSpace: 'normal', lineHeight: 1.35 },
          }}
        >
          {line}
        </ActionButton>
      );
    };

    const pushActionRow = (key: string, line: string, pick: TPageEditableComponentPick): void => {
      out.push(
        <ActionButton
          key={key}
          iconProps={{ iconName: 'BarChartVertical' }}
          onClick={() => onPick(pick)}
          styles={{
            root: {
              width: '100%',
              justifyContent: 'flex-start',
              height: 'auto',
              minHeight: 44,
              padding: '10px 12px',
              borderBottom: '1px solid #edebe9',
            },
            flexContainer: { flexGrow: 1, textAlign: 'left' },
            label: { whiteSpace: 'normal', lineHeight: 1.35 },
          }}
        >
          {line}
        </ActionButton>
      );
    };

    if (config.mode === 'list') {
      const sections = getEffectiveListPageSections(config);
      const blocks = flattenBlocks(sections);

      pushRow('layout', 'Layout da página - Configurar', { kind: 'listLayout' });

      for (const block of blocks) {
        if (block.type === 'dashboard') {
          const dash = resolveDashboardForListBlock(block, rootDash);
          const title = listOrDashTitle(config, block);
          const titlePart = title ? ` «${title}»` : '';
          if (dashboardVisible(dash)) {
            if (dash.dashboardType === 'charts') {
              pushRow(`dash-series-${block.id}`, `Dashboard (gráficos)${titlePart} - Configurar`, {
                kind: 'dashboardSeries',
                blockId: block.id,
              });
            } else {
              pushRow(`dash-cards-${block.id}`, `Dashboard (cartões)${titlePart} - Configurar`, {
                kind: 'dashboardCards',
                blockId: block.id,
              });
              pushActionRow(
                `dash-charts-${block.id}`,
                `Converter dashboard (cartões)${titlePart} para gráficos`,
                { kind: 'dashboardToCharts', blockId: block.id }
              );
            }
          } else {
            pushRow(
              `dash-cfg-${block.id}`,
              `Dashboard (vazio ou desativado)${titlePart} - Configurar`,
              { kind: 'dashboardCards', blockId: block.id }
            );
          }
        }
        if (block.type === 'list') {
          const t = listOrDashTitle(config, block) || config.dataSource.title?.trim() || 'itens';
          pushRow(`tbl-${block.id}`, `Tabela «${t}» - Configurar`, { kind: 'listTable', blockId: block.id });
        }
        if (
          block.type === 'banner' ||
          block.type === 'editor' ||
          block.type === 'sectionTitle' ||
          block.type === 'alert' ||
          block.type === 'buttons'
        ) {
          pushRow(`blk-${block.id}`, `Componente: ${blockContextLabel(config, block)} - Configurar`, {
            kind: 'contentBlock',
            blockId: block.id,
          });
        }
      }

      const hasListBlock = blocks.some((b) => b.type === 'list');
      if (!hasListBlock) {
        const t = config.dataSource.title?.trim() || 'itens';
        pushRow('tbl-fallback', `Tabela «${t}» - Configurar`, {
          kind: 'listTable',
          blockId: LEGACY_LIST_PAGE_LIST_BLOCK_ID,
        });
      }
    }

    if (config.mode === 'formManager') {
      pushRow('fm', 'Formulário - Configurar', { kind: 'formManager' });
    }

    if (config.mode === 'projectManagement') {
      const t = config.dataSource.title?.trim() || 'itens';
      pushRow('pm', `Tabela de projeto «${t}» - Configurar`, { kind: 'projectTable' });
    }

    return out;
  }, [config, onPick, rootDash]);

  return (
    <Modal isOpen={isOpen} onDismiss={onDismiss} isBlocking styles={{ main: { maxWidth: 560, width: '92%' } }}>
      <Stack styles={{ root: { padding: '20px 24px 24px' } }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="xLarge" styles={{ root: { fontWeight: 700 } }}>
            Componentes desta página
          </Text>
          <IconButton iconProps={{ iconName: 'ChromeClose' }} ariaLabel="Fechar" onClick={onDismiss} />
        </Stack>
        <Text variant="small" styles={{ root: { color: '#605e5c', marginTop: 8, marginBottom: 12 } }}>
          Toque na linha do componente para configurar. As alterações ficam na memória até guardar a página.
        </Text>
        <div style={{ maxHeight: '60vh', overflowY: 'auto', border: '1px solid #edebe9', borderRadius: 4 }}>
          <Stack>{rows.length > 0 ? rows : <Text styles={{ root: { padding: 16 } }}>Nada configurável neste modo.</Text>}</Stack>
        </div>
      </Stack>
    </Modal>
  );
};
