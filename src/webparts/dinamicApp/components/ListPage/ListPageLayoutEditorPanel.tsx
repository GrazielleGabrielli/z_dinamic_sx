import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  IconButton,
  Dropdown,
  IDropdownOption,
  Separator,
} from '@fluentui/react';
import type {
  IDashboardConfig,
  IListPageBlock,
  IListPageLayoutConfig,
  IListPageSection,
  TListPageBlockType,
  TListPageSectionLayout,
} from '../../core/config/types';
import {
  defaultAlertConfig,
  defaultBannerConfig,
  defaultRichEditorConfig,
  defaultSectionTitleConfig,
} from '../../core/listPage/listPageBlockConfigUtils';
import {
  cloneDashboardConfig,
  columnCountForLayout,
  countDashboardBlocksInSections,
  normalizeListPageLayoutDashboards,
  reshapeSectionColumns,
} from '../../core/listPage/listPageLayoutUtils';
import { ListPageBlockConfigPanel } from './ListPageBlockConfigPanel';

export interface IListPageLayoutEditorPanelProps {
  isOpen: boolean;
  value: IListPageLayoutConfig;
  rootDashboard: IDashboardConfig;
  onSave: (next: IListPageLayoutConfig) => void;
  onDismiss: () => void;
}

const LAYOUT_OPTIONS: { key: TListPageSectionLayout; label: string }[] = [
  { key: 'one', label: 'Uma coluna' },
  { key: 'two', label: 'Duas colunas' },
  { key: 'three', label: 'Três colunas' },
  { key: 'oneThirdLeft', label: 'Um terço à esquerda' },
  { key: 'oneThirdRight', label: 'Um terço à direita' },
];

const BLOCK_OPTIONS: IDropdownOption[] = [
  { key: 'dashboard', text: 'Dashboard (cards ou gráficos)' },
  { key: 'list', text: 'Tabela / lista' },
  { key: 'banner', text: 'Banner' },
  { key: 'editor', text: 'Editor de conteúdo' },
  { key: 'sectionTitle', text: 'Título de seção' },
  { key: 'alert', text: 'Alerta / aviso' },
];

function blockTypeLabel(type: TListPageBlockType): string {
  if (type === 'dashboard') return 'Dashboard';
  if (type === 'list') return 'Tabela / lista';
  if (type === 'banner') return 'Banner';
  if (type === 'editor') return 'Editor de conteúdo';
  if (type === 'sectionTitle') return 'Título de seção';
  return 'Alerta / aviso';
}

const VALID_SAVE_BLOCK_TYPES: TListPageBlockType[] = [
  'dashboard',
  'list',
  'banner',
  'editor',
  'sectionTitle',
  'alert',
];

function newId(prefix: string): string {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function cloneLayout(v: IListPageLayoutConfig): IListPageLayoutConfig {
  return {
    sections: v.sections.map((s) => ({
      id: s.id,
      layout: s.layout,
      columns: s.columns.map((col) =>
        col.map((b) => {
          const x: IListPageBlock = { ...b };
          if (b.dashboard) x.dashboard = cloneDashboardConfig(b.dashboard);
          if (b.banner) x.banner = { ...b.banner };
          if (b.editor) x.editor = { ...b.editor };
          if (b.sectionTitle) x.sectionTitle = { ...b.sectionTitle };
          if (b.alert) x.alert = { ...b.alert };
          return x;
        })
      ),
    })),
  };
}

export const ListPageLayoutEditorPanel: React.FC<IListPageLayoutEditorPanelProps> = ({
  isOpen,
  value,
  rootDashboard,
  onSave,
  onDismiss,
}) => {
  const [sections, setSections] = useState<IListPageSection[]>(() => value.sections.slice());
  const [blockConfigPath, setBlockConfigPath] = useState<{
    si: number;
    ci: number;
    bi: number;
  } | null>(null);
  const openedRef = useRef(false);

  useEffect(() => {
    if (!isOpen) {
      openedRef.current = false;
      return;
    }
    if (!openedRef.current) {
      openedRef.current = true;
      setSections(cloneLayout(value).sections);
    }
  }, [isOpen, value]);

  const addSection = (): void => {
    setSections((prev) => [
      ...prev,
      {
        id: newId('sec'),
        layout: 'one',
        columns: [[]],
      },
    ]);
  };

  const removeSection = (index: number): void => {
    setSections((prev) => prev.filter((_, i) => i !== index));
  };

  const setSectionLayout = (index: number, layout: TListPageSectionLayout): void => {
    setSections((prev) => {
      const next = prev.slice();
      const cur = next[index];
      if (!cur) return prev;
      next[index] = reshapeSectionColumns({ ...cur, layout: cur.layout }, layout);
      return next;
    });
  };

  const addBlock = (sectionIndex: number, colIndex: number, type: TListPageBlockType): void => {
    setSections((prev) => {
      const next = prev.slice();
      const sec = next[sectionIndex];
      if (!sec || !sec.columns[colIndex]) return prev;
      const cols = sec.columns.map((c) => c.slice());
      const dashboardsBefore = countDashboardBlocksInSections(prev);
      const block: IListPageBlock =
        type === 'dashboard'
          ? dashboardsBefore >= 1
            ? { id: newId('blk'), type, dashboard: cloneDashboardConfig(rootDashboard) }
            : { id: newId('blk'), type }
          : type === 'banner'
          ? { id: newId('blk'), type: 'banner', banner: defaultBannerConfig() }
          : type === 'editor'
          ? { id: newId('blk'), type: 'editor', editor: defaultRichEditorConfig() }
          : type === 'sectionTitle'
          ? { id: newId('blk'), type: 'sectionTitle', sectionTitle: defaultSectionTitleConfig() }
          : type === 'alert'
          ? { id: newId('blk'), type: 'alert', alert: defaultAlertConfig() }
          : { id: newId('blk'), type: 'list' };
      cols[colIndex] = [...cols[colIndex], block];
      next[sectionIndex] = { ...sec, columns: cols };
      return normalizeListPageLayoutDashboards({ sections: next }, rootDashboard).sections;
    });
  };

  const removeBlock = (sectionIndex: number, colIndex: number, blockIndex: number): void => {
    setSections((prev) => {
      const next = prev.slice();
      const sec = next[sectionIndex];
      if (!sec) return prev;
      const cols = sec.columns.map((c) => c.slice());
      if (!cols[colIndex]) return prev;
      cols[colIndex] = cols[colIndex].filter((_, bi) => bi !== blockIndex);
      next[sectionIndex] = { ...sec, columns: cols };
      return next;
    });
  };

  const handleSave = (): void => {
    const cleaned: IListPageSection[] = [];
    for (let i = 0; i < sections.length; i++) {
      const s = sections[i];
      const nc = columnCountForLayout(s.layout);
      const cols: IListPageBlock[][] = [];
      for (let c = 0; c < nc; c++) {
        cols.push(
          (s.columns[c] ?? []).filter((b) => b.id && VALID_SAVE_BLOCK_TYPES.indexOf(b.type) !== -1)
        );
      }
      cleaned.push({
        id: s.id.trim() || newId('sec'),
        layout: s.layout,
        columns: cols,
      });
    }
    if (cleaned.length === 0) {
      cleaned.push({ id: newId('sec'), layout: 'one', columns: [[{ id: newId('blk'), type: 'list' }]] });
    }
    let hasList = false;
    for (let si = 0; si < cleaned.length && !hasList; si++) {
      const cols = cleaned[si].columns;
      for (let ci = 0; ci < cols.length && !hasList; ci++) {
        const blocks = cols[ci];
        for (let bi = 0; bi < blocks.length; bi++) {
          if (blocks[bi].type === 'list') {
            hasList = true;
            break;
          }
        }
      }
    }
    if (!hasList) {
      cleaned.push({ id: newId('sec'), layout: 'one', columns: [[{ id: newId('blk'), type: 'list' }]] });
    }
    const normalized = normalizeListPageLayoutDashboards({ sections: cleaned }, rootDashboard).sections;
    onSave({ sections: normalized });
    onDismiss();
  };

  const editingBlock =
    blockConfigPath !== null
      ? sections[blockConfigPath.si]?.columns[blockConfigPath.ci]?.[blockConfigPath.bi] ?? null
      : null;

  const applyBlockConfig = (next: IListPageBlock): void => {
    if (!blockConfigPath) return;
    const { si, ci, bi } = blockConfigPath;
    setSections((prev) => {
      const copy = prev.slice();
      const sec = copy[si];
      if (!sec) return prev;
      const cols = sec.columns.map((c) => c.slice());
      const col = cols[ci];
      if (!col || col[bi] === undefined) return prev;
      const nextCol = col.slice();
      nextCol[bi] = next;
      cols[ci] = nextCol;
      copy[si] = { ...sec, columns: cols };
      return copy;
    });
    setBlockConfigPath(null);
  };

  return (
    <>
    <Panel
      isOpen={isOpen}
      type={PanelType.custom}
      customWidth="520px"
      headerText="Layout da página (modo lista)"
      onDismiss={onDismiss}
      closeButtonAriaLabel="Fechar"
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="Salvar" onClick={handleSave} />
          <DefaultButton text="Cancelar" onClick={onDismiss} />
        </Stack>
      )}
    >
      <Stack tokens={{ childrenGap: 16 }} styles={{ root: { paddingTop: 12, maxWidth: '100%' } }}>
        <Text variant="small" styles={{ root: { color: '#605e5c', lineHeight: 1.5 } }}>
          Monte seções como em páginas modernas: colunas por seção e blocos (dashboard, tabela, banner, editor,
          título de seção ou alerta). Use a engrenagem para configurar esses blocos de conteúdo. Lista e
          dashboard usam os botões da barra da página.
        </Text>
        {sections.map((sec, si) => (
          <Stack
            key={sec.id}
            tokens={{ childrenGap: 12 }}
            styles={{
              root: {
                padding: 14,
                border: '1px solid #edebe9',
                borderRadius: 8,
                background: '#faf9f8',
                maxWidth: '100%',
              },
            }}
          >
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                Seção {si + 1}
              </Text>
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title="Remover seção"
                ariaLabel="Remover seção"
                onClick={() => removeSection(si)}
                disabled={sections.length <= 1}
              />
            </Stack>
            <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
              Colunas
            </Text>
            <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
              {LAYOUT_OPTIONS.map((opt) => {
                const selected = sec.layout === opt.key;
                return (
                  <DefaultButton
                    key={opt.key}
                    text={opt.label}
                    primary={selected}
                    onClick={() => setSectionLayout(si, opt.key)}
                    styles={{ root: { minHeight: 36, maxWidth: '100%' } }}
                  />
                );
              })}
            </Stack>
            <Separator />
            {sec.columns.map((blocks, ci) => (
              <Stack key={`${sec.id}_col_${ci}`} tokens={{ childrenGap: 8 }} styles={{ root: { maxWidth: '100%' } }}>
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  Coluna {ci + 1}
                </Text>
                {blocks.map((b, bi) => (
                  <Stack
                    key={b.id}
                    horizontal
                    verticalAlign="center"
                    horizontalAlign="space-between"
                    styles={{
                      root: {
                        padding: '8px 10px',
                        background: '#fff',
                        borderRadius: 4,
                        border: '1px solid #edebe9',
                      },
                    }}
                  >
                    <Text variant="small">{blockTypeLabel(b.type)}</Text>
                    <Stack horizontal verticalAlign="center">
                      {(b.type === 'banner' ||
                        b.type === 'editor' ||
                        b.type === 'sectionTitle' ||
                        b.type === 'alert') && (
                        <IconButton
                          iconProps={{ iconName: 'Settings' }}
                          title="Configurar bloco"
                          ariaLabel="Configurar bloco"
                          onClick={() => setBlockConfigPath({ si, ci, bi })}
                        />
                      )}
                      <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title="Remover bloco"
                        ariaLabel="Remover bloco"
                        onClick={() => removeBlock(si, ci, bi)}
                      />
                    </Stack>
                  </Stack>
                ))}
                <Dropdown
                  placeholder="Adicionar bloco nesta coluna…"
                  options={BLOCK_OPTIONS}
                  onChange={(_, opt) => {
                    if (opt) addBlock(si, ci, String(opt.key) as TListPageBlockType);
                  }}
                  styles={{ root: { maxWidth: '100%' } }}
                />
              </Stack>
            ))}
          </Stack>
        ))}
        <DefaultButton text="Adicionar seção" iconProps={{ iconName: 'Add' }} onClick={addSection} />
      </Stack>
    </Panel>
    <ListPageBlockConfigPanel
      isOpen={blockConfigPath !== null && editingBlock !== null}
      block={editingBlock}
      onDismiss={() => setBlockConfigPath(null)}
      onApply={applyBlockConfig}
    />
    </>
  );
};
