import * as React from 'react';
import {
  ActionButton,
  IconButton,
  Modal,
  ScrollablePane,
  ScrollbarVisibility,
  Stack,
  Text,
} from '@fluentui/react';
import type { IDashboardConfig, IDynamicViewConfig, IListPageBlock, IListPageSection } from '../core/config/types';
import {
  effectiveConfigForListPageBlock,
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
  listSections: IListPageSection[];
  onPick: (pick: TPageEditableComponentPick) => void;
}

function flattenBlocks(sections: IListPageSection[]): IListPageBlock[] {
  const out: IListPageBlock[] = [];
  for (const sec of sections) {
    for (const col of sec.columns) {
      for (const b of col) {
        out.push(b);
      }
    }
  }
  return out;
}

function dashboardVisible(d: IDashboardConfig): boolean {
  return d.enabled && (d.dashboardType === 'charts' || (d.cardsCount ?? 0) > 0);
}

function blockContextLabel(config: IDynamicViewConfig, block: IListPageBlock): string {
  switch (block.type) {
    case 'dashboard':
      return 'Dashboard';
    case 'list': {
      const t = effectiveConfigForListPageBlock(config, block).dataSource.title?.trim();
      return t ? `Lista — ${t}` : 'Lista';
    }
    case 'banner': {
      const t = block.banner?.title?.trim();
      return t ? `Banner — ${t}` : 'Banner';
    }
    case 'editor': {
      const t = block.editor?.title?.trim();
      return t ? `Texto — ${t}` : 'Bloco de texto';
    }
    case 'sectionTitle': {
      const t = block.sectionTitle?.title?.trim();
      return t ? `Título — ${t}` : 'Título de secção';
    }
    case 'alert': {
      const t = block.alert?.title?.trim();
      return t ? `Alerta — ${t}` : 'Alerta';
    }
    case 'buttons': {
      const first = block.buttons?.items?.[0]?.label?.trim();
      return first ? `Botões — ${first}` : 'Botões';
    }
    default:
      return block.type;
  }
}

export const PageEditableComponentsModal: React.FC<IPageEditableComponentsModalProps> = ({
  isOpen,
  onDismiss,
  config,
  listSections,
  onPick,
}) => {
  const rootDash = config.dashboard;
  const blocks = React.useMemo(() => flattenBlocks(listSections), [listSections]);

  const rows = React.useMemo(() => {
    const out: React.ReactNode[] = [];

    const pushRow = (key: string, label: string, sub: string | undefined, pick: TPageEditableComponentPick): void => {
      out.push(
        <Stack key={key} tokens={{ childrenGap: 2 }} styles={{ root: { padding: '4px 0' } }}>
          <ActionButton
            iconProps={{ iconName: 'Edit' }}
            onClick={() => onPick(pick)}
            styles={{ root: { height: 'auto', minHeight: 36, padding: '8px 10px' } }}
          >
            <Stack tokens={{ childrenGap: 2 }}>
              <Text styles={{ root: { fontWeight: 600 } }}>{label}</Text>
              {sub ? (
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  {sub}
                </Text>
              ) : null}
            </Stack>
          </ActionButton>
        </Stack>
      );
    };

    if (config.mode === 'list') {
      pushRow('layout', 'Layout da página', 'Secções, colunas e blocos da vista em lista', { kind: 'listLayout' });
      for (const block of blocks) {
        if (block.type === 'dashboard') {
          const dash = resolveDashboardForListBlock(block, rootDash);
          const ctx = blockContextLabel(config, block);
          if (dashboardVisible(dash)) {
            if (dash.dashboardType === 'charts') {
              pushRow(
                `dash-series-${block.id}`,
                'Gráficos e séries',
                ctx,
                { kind: 'dashboardSeries', blockId: block.id }
              );
            } else {
              pushRow(
                `dash-cards-${block.id}`,
                'Cartões do dashboard',
                ctx,
                { kind: 'dashboardCards', blockId: block.id }
              );
              pushRow(
                `dash-charts-${block.id}`,
                'Converter dashboard para gráficos',
                ctx,
                { kind: 'dashboardToCharts', blockId: block.id }
              );
            }
          } else {
            pushRow(
              `dash-cfg-${block.id}`,
              'Dashboard (desativado ou vazio)',
              `${ctx} — abrir editor de cartões`,
              { kind: 'dashboardCards', blockId: block.id }
            );
          }
        }
        if (block.type === 'list') {
          pushRow(
            `tbl-${block.id}`,
            'Colunas, filtros e PDF',
            blockContextLabel(config, block),
            { kind: 'listTable', blockId: block.id }
          );
        }
        if (
          block.type === 'banner' ||
          block.type === 'editor' ||
          block.type === 'sectionTitle' ||
          block.type === 'alert' ||
          block.type === 'buttons'
        ) {
          pushRow(
            `blk-${block.id}`,
            'Configurar bloco',
            blockContextLabel(config, block),
            { kind: 'contentBlock', blockId: block.id }
          );
        }
      }
    }

    if (config.mode === 'formManager') {
      pushRow('fm', 'Formulário FlexView', 'Passos, campos, regras e anexos', { kind: 'formManager' });
    }

    if (config.mode === 'projectManagement') {
      const t = config.dataSource.title?.trim();
      pushRow('pm', 'Tabela de projeto', t ? `Lista: ${t}` : undefined, { kind: 'projectTable' });
    }

    return out;
  }, [blocks, config, onPick, rootDash]);

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
          Escolha o que pretende editar. As alterações ficam na memória até guardar a página.
        </Text>
        <ScrollablePane
          scrollbarVisibility={ScrollbarVisibility.auto}
          styles={{ root: { maxHeight: 'min(420px, 60vh)', position: 'relative' } }}
        >
          <Stack tokens={{ childrenGap: 4 }}>{rows.length > 0 ? rows : <Text>Nada configurável neste modo.</Text>}</Stack>
        </ScrollablePane>
      </Stack>
    </Modal>
  );
};
