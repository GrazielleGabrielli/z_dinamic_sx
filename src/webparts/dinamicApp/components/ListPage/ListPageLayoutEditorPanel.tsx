import * as React from 'react';
import { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  TextField,
  Link,
  MessageBar,
  MessageBarType,
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
  defaultButtonsConfig,
  defaultRichEditorConfig,
  defaultSectionTitleConfig,
} from '../../core/listPage/listPageBlockConfigUtils';
import {
  cloneDashboardConfig,
  columnCountForLayout,
  countDashboardBlocksInSections,
  mergeExtraListPageColumnsIntoLayout,
  normalizeListPageLayoutDashboards,
  reshapeSectionColumns,
  sanitizeListPageContentPadding,
  sanitizeListPageLayout,
} from '../../core/listPage/listPageLayoutUtils';
import { ListPageBlockConfigPanel } from './ListPageBlockConfigPanel';

export interface IListPageLayoutEditorPanelProps {
  isOpen: boolean;
  value: IListPageLayoutConfig;
  rootDashboard: IDashboardConfig;
  /** Título da lista da vista (regras de contagem no alerta). */
  sourceListTitle?: string;
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
  { key: 'buttons', text: 'Botões' },
];

function blockTypeLabel(type: TListPageBlockType): string {
  if (type === 'dashboard') return 'Dashboard';
  if (type === 'list') return 'Tabela / lista';
  if (type === 'banner') return 'Banner';
  if (type === 'editor') return 'Editor de conteúdo';
  if (type === 'sectionTitle') return 'Título de seção';
  if (type === 'buttons') return 'Botões';
  return 'Alerta / aviso';
}

const VALID_SAVE_BLOCK_TYPES: TListPageBlockType[] = [
  'dashboard',
  'list',
  'banner',
  'editor',
  'sectionTitle',
  'alert',
  'buttons',
];

function newId(prefix: string): string {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function layoutLabel(layout: TListPageSectionLayout): string {
  const f = LAYOUT_OPTIONS.find((o) => o.key === layout);
  return f?.label ?? layout;
}

function ListPageLayoutCollapse(props: {
  title: string;
  isOpen: boolean;
  onToggle: () => void;
  children: React.ReactNode;
  trailing?: React.ReactNode;
}): JSX.Element {
  return (
    <Stack
      styles={{
        root: {
          border: '1px solid #edebe9',
          borderRadius: 10,
          background: '#ffffff',
          boxShadow: '0 1px 2px rgba(0,0,0,0.04)',
          overflow: 'hidden',
          maxWidth: '100%',
          minWidth: 0,
          width: '100%',
          boxSizing: 'border-box',
        },
      }}
    >
      <Stack
        horizontal
        verticalAlign="center"
        tokens={{ childrenGap: 2 }}
        styles={{
          root: {
            padding: '10px 12px',
            background: props.isOpen ? '#faf9f8' : '#ffffff',
            borderBottom: props.isOpen ? '1px solid #edebe9' : undefined,
            userSelect: 'none',
          },
        }}
      >
        <IconButton
          iconProps={{ iconName: props.isOpen ? 'ChevronDown' : 'ChevronRight' }}
          title={props.isOpen ? 'Recolher' : 'Expandir'}
          aria-expanded={props.isOpen}
          onClick={(e) => {
            e.preventDefault();
            props.onToggle();
          }}
          styles={{ root: { width: 32, height: 32 } }}
        />
        <Text
          variant="smallPlus"
          styles={{ root: { fontWeight: 600, cursor: 'pointer', flex: 1, color: '#323130' } }}
          onClick={props.onToggle}
        >
          {props.title}
        </Text>
        {props.trailing ?? null}
      </Stack>
      {props.isOpen ? (
        <div
          style={{
            padding: '14px 14px 16px 18px',
            maxWidth: '100%',
            minWidth: 0,
            width: '100%',
            boxSizing: 'border-box',
            display: 'flex',
            flexDirection: 'column',
            gap: 12,
          }}
        >
          {props.children}
        </div>
      ) : null}
    </Stack>
  );
}

function finalizeListPageLayoutForSave(
  sections: IListPageSection[],
  rootDashboard: IDashboardConfig,
  contentPaddingInput: string
): IListPageLayoutConfig {
  const cleaned: IListPageSection[] = [];
  for (let i = 0; i < sections.length; i++) {
    const s = sections[i];
    const merged = mergeExtraListPageColumnsIntoLayout(s.layout, s.columns.map((col) => col.slice()));
    const nc = columnCountForLayout(s.layout);
    const cols: IListPageBlock[][] = [];
    for (let c = 0; c < nc; c++) {
      cols.push(
        (merged[c] ?? []).filter((b) => b.id && VALID_SAVE_BLOCK_TYPES.indexOf(b.type) !== -1)
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
  const pad = sanitizeListPageContentPadding(contentPaddingInput);
  const base: IListPageLayoutConfig = {
    sections: cleaned,
    ...(pad ? { contentPadding: pad } : {}),
  };
  return normalizeListPageLayoutDashboards(base, rootDashboard);
}

function cloneLayout(v: IListPageLayoutConfig): IListPageLayoutConfig {
  const pad = v.contentPadding?.trim();
  return {
    sections: v.sections.map((s) => {
      const columnsMerged = mergeExtraListPageColumnsIntoLayout(s.layout, s.columns.map((col) => col.slice()));
      return {
      id: s.id,
      layout: s.layout,
      columns: columnsMerged.map((col) =>
        col.map((b) => {
          const x: IListPageBlock = { ...b };
          if (b.dashboard) x.dashboard = cloneDashboardConfig(b.dashboard);
          if (b.banner) x.banner = { ...b.banner };
          if (b.editor) x.editor = { ...b.editor };
          if (b.sectionTitle) x.sectionTitle = { ...b.sectionTitle };
          if (b.alert) {
            x.alert = { ...b.alert };
            if (b.alert.countRules?.length) {
              x.alert.countRules = b.alert.countRules.map((r) => ({ ...r }));
            }
          }
          if (b.buttons) {
            x.buttons = { items: (b.buttons.items ?? []).map((it) => ({ ...it })) };
          }
          return x;
        })
      ),
    };
    }),
    ...(pad ? { contentPadding: pad } : {}),
  };
}

export const ListPageLayoutEditorPanel: React.FC<IListPageLayoutEditorPanelProps> = ({
  isOpen,
  value,
  rootDashboard,
  sourceListTitle,
  onSave,
  onDismiss,
}) => {
  const [sections, setSections] = useState<IListPageSection[]>(() => value.sections.slice());
  const [jsonOpen, setJsonOpen] = useState(false);
  const [jsonPanelText, setJsonPanelText] = useState('');
  const [jsonPanelErr, setJsonPanelErr] = useState<string | undefined>(undefined);
  const [blockConfigPath, setBlockConfigPath] = useState<{
    si: number;
    ci: number;
    bi: number;
  } | null>(null);
  const [helpOpen, setHelpOpen] = useState(false);
  const [paddingOpen, setPaddingOpen] = useState(false);
  const [contentPadding, setContentPadding] = useState('');
  const [sectionOpen, setSectionOpen] = useState<Record<string, boolean>>({});
  const openedRef = useRef(false);
  const contentPaddingRef = useRef(contentPadding);
  contentPaddingRef.current = contentPadding;

  useEffect(() => {
    if (!isOpen) {
      openedRef.current = false;
      return;
    }
    if (!openedRef.current) {
      openedRef.current = true;
      const nextSections = cloneLayout(value).sections;
      setSections(nextSections);
      setContentPadding(value.contentPadding ?? '');
      setSectionOpen({});
      setHelpOpen(false);
    }
  }, [isOpen, value]);

  const layoutJsonPreview = useMemo(
    () => JSON.stringify(finalizeListPageLayoutForSave(sections, rootDashboard, contentPadding), null, 2),
    [sections, rootDashboard, contentPadding]
  );
  const layoutJsonPreviewRef = useRef(layoutJsonPreview);
  layoutJsonPreviewRef.current = layoutJsonPreview;
  useEffect(() => {
    if (jsonOpen) {
      setJsonPanelText(layoutJsonPreviewRef.current);
      setJsonPanelErr(undefined);
    }
  }, [jsonOpen]);

  const applyLayoutJsonFromPanel = useCallback(() => {
    setJsonPanelErr(undefined);
    try {
      const parsed: unknown = JSON.parse(jsonPanelText);
      const sanitized = sanitizeListPageLayout(parsed);
      if (!sanitized) {
        setJsonPanelErr('JSON inválido ou estrutura não reconhecida.');
        return;
      }
      const next = finalizeListPageLayoutForSave(
        sanitized.sections,
        rootDashboard,
        sanitized.contentPadding ?? ''
      );
      setSections(next.sections);
      setContentPadding(next.contentPadding ?? '');
      setSectionOpen({});
      setJsonPanelText(JSON.stringify(next, null, 2));
    } catch (e) {
      setJsonPanelErr(e instanceof Error ? e.message : String(e));
    }
  }, [jsonPanelText, rootDashboard]);

  const addSection = (): void => {
    const id = newId('sec');
    setSections((prev) => [...prev, { id, layout: 'one', columns: [[]] }]);
  };

  const removeSection = (index: number): void => {
    const sid = sections[index]?.id;
    setSections((prev) => prev.filter((_, i) => i !== index));
    if (sid) {
      setSectionOpen((p) => {
        const n = { ...p };
        delete n[sid];
        return n;
      });
    }
  };

  const moveSection = (index: number, delta: -1 | 1): void => {
    setSections((prev) => {
      const j = index + delta;
      if (j < 0 || j >= prev.length) return prev;
      const next = prev.slice();
      const t = next[index];
      next[index] = next[j];
      next[j] = t;
      return next;
    });
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
          : type === 'buttons'
          ? { id: newId('blk'), type: 'buttons', buttons: defaultButtonsConfig() }
          : { id: newId('blk'), type: 'list' };
      cols[colIndex] = [...cols[colIndex], block];
      next[sectionIndex] = { ...sec, columns: cols };
      const pad = sanitizeListPageContentPadding(contentPaddingRef.current);
      return normalizeListPageLayoutDashboards(
        { sections: next, ...(pad ? { contentPadding: pad } : {}) },
        rootDashboard
      ).sections;
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

  const moveBlockInColumn = (
    sectionIndex: number,
    colIndex: number,
    blockIndex: number,
    delta: -1 | 1
  ): void => {
    setSections((prev) => {
      const next = prev.slice();
      const sec = next[sectionIndex];
      if (!sec) return prev;
      const cols = sec.columns.map((c) => c.slice());
      const col = cols[colIndex];
      if (!col) return prev;
      const j = blockIndex + delta;
      if (j < 0 || j >= col.length) return prev;
      const newCol = col.slice();
      const t = newCol[blockIndex];
      newCol[blockIndex] = newCol[j];
      newCol[j] = t;
      cols[colIndex] = newCol;
      next[sectionIndex] = { ...sec, columns: cols };
      return next;
    });
  };

  const handleSave = (): void => {
    onSave(finalizeListPageLayoutForSave(sections, rootDashboard, contentPadding));
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
      <Stack
        tokens={{ childrenGap: 12 }}
        styles={{
          root: {
            paddingTop: 12,
            maxWidth: '100%',
            width: '100%',
            boxSizing: 'border-box',
          },
        }}
      >
        <ListPageLayoutCollapse title="Ajuda" isOpen={helpOpen} onToggle={() => setHelpOpen((v) => !v)}>
          <div style={{ display: 'flex', justifyContent: 'flex-end', width: '100%', minWidth: 0 }}>
            <Link onClick={() => setJsonOpen(true)}>JSON (ver / colar)</Link>
          </div>
          <p
            style={{
              margin: 0,
              minWidth: 0,
              width: '100%',
              color: '#605e5c',
              fontSize: 12,
              lineHeight: 1.65,
              whiteSpace: 'normal',
              wordBreak: 'break-word',
              overflowWrap: 'anywhere',
            }}
          >
            Monte seções como em páginas modernas: colunas por seção e blocos (dashboard, tabela, banner, editor,
            título de seção, alerta ou botões). Use as setas no cabeçalho de cada seção para alterar a ordem (qual
            aparece primeiro na página). Use a engrenagem para configurar esses blocos de conteúdo. Lista e dashboard
            usam os botões da barra da página.
          </p>
        </ListPageLayoutCollapse>

        <ListPageLayoutCollapse title="Espaçamento" isOpen={paddingOpen} onToggle={() => setPaddingOpen((v) => !v)}>
          <TextField
            label="Padding da área do layout"
            value={contentPadding}
            onChange={(_, v) => setContentPadding(v ?? '')}
            placeholder="ex.: 16px 24px"
            description="Formato CSS: vertical e horizontal (ex. 16px 24px). Opcional: até 4 valores «Npx» separados por espaços."
          />
        </ListPageLayoutCollapse>

        {sections.map((sec, si) => (
          <ListPageLayoutCollapse
            key={sec.id}
            title={`Seção ${si + 1} · ${layoutLabel(sec.layout)}`}
            isOpen={sectionOpen[sec.id] === true}
            onToggle={() =>
              setSectionOpen((p) => {
                const wasOpen = p[sec.id] === true;
                return { ...p, [sec.id]: wasOpen ? false : true };
              })
            }
            trailing={
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 0 }}>
                <IconButton
                  iconProps={{ iconName: 'ChevronUp' }}
                  title="Mover seção para cima"
                  ariaLabel="Mover seção para cima"
                  disabled={si === 0}
                  onClick={(e) => {
                    e.stopPropagation();
                    moveSection(si, -1);
                  }}
                  styles={{ root: { width: 32, height: 32 } }}
                />
                <IconButton
                  iconProps={{ iconName: 'ChevronDown' }}
                  title="Mover seção para baixo"
                  ariaLabel="Mover seção para baixo"
                  disabled={si >= sections.length - 1}
                  onClick={(e) => {
                    e.stopPropagation();
                    moveSection(si, 1);
                  }}
                  styles={{ root: { width: 32, height: 32 } }}
                />
                <IconButton
                  iconProps={{ iconName: 'Delete' }}
                  title="Remover seção"
                  ariaLabel="Remover seção"
                  onClick={(e) => {
                    e.stopPropagation();
                    removeSection(si);
                  }}
                  disabled={sections.length <= 1}
                  styles={{ root: { width: 32, height: 32 } }}
                />
              </Stack>
            }
          >
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
                      <IconButton
                        iconProps={{ iconName: 'ChevronUp' }}
                        title="Subir na coluna"
                        ariaLabel="Subir na coluna"
                        disabled={bi === 0}
                        onClick={() => moveBlockInColumn(si, ci, bi, -1)}
                        styles={{ root: { width: 28, height: 28 } }}
                      />
                      <IconButton
                        iconProps={{ iconName: 'ChevronDown' }}
                        title="Descer na coluna"
                        ariaLabel="Descer na coluna"
                        disabled={bi >= blocks.length - 1}
                        onClick={() => moveBlockInColumn(si, ci, bi, 1)}
                        styles={{ root: { width: 28, height: 28 } }}
                      />
                      {(b.type === 'banner' ||
                        b.type === 'editor' ||
                        b.type === 'sectionTitle' ||
                        b.type === 'alert' ||
                        b.type === 'buttons') && (
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
                  key={`addblk_${sec.id}_${ci}_${blocks.length}`}
                  placeholder="Adicionar bloco nesta coluna…"
                  options={BLOCK_OPTIONS}
                  onChange={(_, opt) => {
                    if (opt) addBlock(si, ci, String(opt.key) as TListPageBlockType);
                  }}
                  styles={{ root: { maxWidth: '100%' } }}
                />
              </Stack>
            ))}
          </ListPageLayoutCollapse>
        ))}
        <DefaultButton text="Adicionar seção" iconProps={{ iconName: 'Add' }} onClick={addSection} />
      </Stack>
    </Panel>
    <ListPageBlockConfigPanel
      isOpen={blockConfigPath !== null && editingBlock !== null}
      block={editingBlock}
      listTitle={sourceListTitle ?? ''}
      onDismiss={() => setBlockConfigPath(null)}
      onApply={applyBlockConfig}
    />
    <Panel
      isOpen={jsonOpen}
      type={PanelType.medium}
      headerText="Layout da página (JSON)"
      onDismiss={() => setJsonOpen(false)}
    >
      <Text
        variant="small"
        styles={{
          root: {
            display: 'block',
            width: '100%',
            maxWidth: '100%',
            color: '#605e5c',
            lineHeight: 1.55,
            whiteSpace: 'normal',
            wordWrap: 'break-word',
            overflowWrap: 'break-word',
            marginBottom: 10,
          },
        }}
      >
        Objeto com «sections» (listPageLayout). Aplicar carrega no editor; Salvar na janela principal grava na vista.
      </Text>
      {jsonPanelErr && (
        <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 8 } }}>
          {jsonPanelErr}
        </MessageBar>
      )}
      <TextField
        multiline
        rows={22}
        value={jsonPanelText}
        onChange={(_, v) => setJsonPanelText(v ?? '')}
        styles={{ root: { fontFamily: 'monospace', fontSize: 12 } }}
      />
      <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 12 } }}>
        <PrimaryButton text="Aplicar JSON" onClick={() => applyLayoutJsonFromPanel()} />
        <DefaultButton text="Fechar" onClick={() => setJsonOpen(false)} />
      </Stack>
    </Panel>
    </>
  );
};
