import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  TextField,
  Toggle,
  Dropdown,
  IDropdownOption,
  Slider,
  PrimaryButton,
  DefaultButton,
} from '@fluentui/react';
import type {
  IListPageAlertBlockConfig,
  IListPageBannerBlockConfig,
  IListPageBlock,
  IListPageRichEditorBlockConfig,
  IListPageSectionTitleBlockConfig,
  TListPageAlertVariant,
  TListPageBannerContentAlign,
  TListPageSectionTitleSize,
} from '../../core/config/types';
import {
  defaultAlertConfig,
  defaultBannerConfig,
  defaultRichEditorConfig,
  defaultSectionTitleConfig,
} from '../../core/listPage/listPageBlockConfigUtils';
import { ListPageRichQuillEditor } from './ListPageRichQuillEditor';

export interface IListPageBlockConfigPanelProps {
  isOpen: boolean;
  block: IListPageBlock | null;
  onDismiss: () => void;
  onApply: (next: IListPageBlock) => void;
}

const ALIGN_OPTIONS: IDropdownOption[] = [
  { key: 'left', text: 'Esquerda' },
  { key: 'center', text: 'Centro' },
  { key: 'right', text: 'Direita' },
];

const SECTION_TITLE_SIZE_OPTIONS: IDropdownOption[] = [
  { key: 'sm', text: 'Pequeno' },
  { key: 'md', text: 'Médio' },
  { key: 'lg', text: 'Grande' },
];

const ALERT_VARIANT_OPTIONS: IDropdownOption[] = [
  { key: 'info', text: 'Informação' },
  { key: 'success', text: 'Sucesso' },
  { key: 'warning', text: 'Atenção' },
  { key: 'error', text: 'Erro' },
];

function panelHeaderForBlockType(t: IListPageBlock['type']): string {
  if (t === 'banner') return 'Banner';
  if (t === 'editor') return 'Editor de conteúdo';
  if (t === 'sectionTitle') return 'Título de seção';
  if (t === 'alert') return 'Alerta / aviso';
  return '';
}

function panelWidthForBlockType(t: IListPageBlock['type']): string {
  return t === 'editor' ? '620px' : '480px';
}

export const ListPageBlockConfigPanel: React.FC<IListPageBlockConfigPanelProps> = ({
  isOpen,
  block,
  onDismiss,
  onApply,
}) => {
  const [banner, setBanner] = useState<IListPageBannerBlockConfig>(defaultBannerConfig);
  const [editor, setEditor] = useState<IListPageRichEditorBlockConfig>(defaultRichEditorConfig);
  const [editorHtmlSourceMode, setEditorHtmlSourceMode] = useState(false);
  const [sectionTitle, setSectionTitle] = useState<IListPageSectionTitleBlockConfig>(defaultSectionTitleConfig);
  const [alertCfg, setAlertCfg] = useState<IListPageAlertBlockConfig>(defaultAlertConfig);

  useEffect(() => {
    if (!isOpen || !block) return;
    if (block.type === 'banner') {
      setBanner(block.banner ? { ...block.banner } : defaultBannerConfig());
    }
    if (block.type === 'editor') {
      setEditor(block.editor ? { ...block.editor } : defaultRichEditorConfig());
      setEditorHtmlSourceMode(false);
    }
    if (block.type === 'sectionTitle') {
      setSectionTitle(block.sectionTitle ? { ...block.sectionTitle } : defaultSectionTitleConfig());
    }
    if (block.type === 'alert') {
      setAlertCfg(block.alert ? { ...block.alert } : defaultAlertConfig());
    }
  }, [isOpen, block]);

  if (
    !block ||
    (block.type !== 'banner' &&
      block.type !== 'editor' &&
      block.type !== 'sectionTitle' &&
      block.type !== 'alert')
  ) {
    return null;
  }

  const handleSave = (): void => {
    if (block.type === 'banner') {
      onApply({ ...block, banner });
    } else if (block.type === 'editor') {
      onApply({ ...block, editor: { ...editor, html: editor.html } });
    } else if (block.type === 'sectionTitle') {
      onApply({ ...block, sectionTitle });
    } else {
      onApply({ ...block, alert: alertCfg });
    }
  };

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.custom}
      customWidth={panelWidthForBlockType(block.type)}
      headerText={panelHeaderForBlockType(block.type)}
      onDismiss={onDismiss}
      closeButtonAriaLabel="Fechar"
      isLightDismiss
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="Aplicar" onClick={handleSave} />
          <DefaultButton text="Cancelar" onClick={onDismiss} />
        </Stack>
      )}
    >
      <Stack tokens={{ childrenGap: 14 }} styles={{ root: { paddingTop: 8, maxWidth: '100%' } }}>
        {block.type === 'banner' ? (
          <>
            <TextField
              label="URL da imagem"
              value={banner.imageUrl}
              onChange={(_, v) => setBanner((b) => ({ ...b, imageUrl: v ?? '' }))}
            />
            <TextField
              label="Título"
              value={banner.title}
              onChange={(_, v) => setBanner((b) => ({ ...b, title: v ?? '' }))}
            />
            <TextField
              label="Subtítulo"
              multiline
              rows={2}
              value={banner.subtitle}
              onChange={(_, v) => setBanner((b) => ({ ...b, subtitle: v ?? '' }))}
            />
            <TextField
              label="Link de destino"
              value={banner.linkUrl}
              onChange={(_, v) => setBanner((b) => ({ ...b, linkUrl: v ?? '' }))}
            />
            <Toggle
              label="Abrir em nova aba"
              checked={banner.openInNewTab}
              onChange={(_, c) => setBanner((b) => ({ ...b, openInNewTab: Boolean(c) }))}
            />
            <TextField
              label="Alt da imagem"
              value={banner.imageAlt}
              onChange={(_, v) => setBanner((b) => ({ ...b, imageAlt: v ?? '' }))}
            />
            <Dropdown
              label="Alinhamento do conteúdo"
              selectedKey={banner.contentAlign}
              options={ALIGN_OPTIONS}
              onChange={(_, opt) => {
                const k = opt?.key as TListPageBannerContentAlign | undefined;
                if (k === 'left' || k === 'center' || k === 'right') {
                  setBanner((b) => ({ ...b, contentAlign: k }));
                }
              }}
            />
            <TextField
              label="Altura do banner (px)"
              type="number"
              value={String(banner.heightPx)}
              onChange={(_, v) => {
                const n = parseInt(String(v ?? ''), 10);
                if (isNaN(n)) return;
                setBanner((b) => ({ ...b, heightPx: Math.min(800, Math.max(80, n)) }));
              }}
            />
            <Stack tokens={{ childrenGap: 6 }}>
              <Text variant="small">Overlay / escurecimento ({Math.round(banner.overlayOpacity * 100)}%)</Text>
              <Slider
                min={0}
                max={100}
                step={5}
                value={banner.overlayOpacity * 100}
                showValue
                onChange={(n) => setBanner((b) => ({ ...b, overlayOpacity: n / 100 }))}
              />
            </Stack>
            <Toggle
              label="Exibir botão"
              checked={banner.showButton}
              onChange={(_, c) => setBanner((b) => ({ ...b, showButton: Boolean(c) }))}
            />
            <TextField
              label="Texto do botão"
              disabled={!banner.showButton}
              value={banner.buttonText}
              onChange={(_, v) => setBanner((b) => ({ ...b, buttonText: v ?? '' }))}
            />
          </>
        ) : null}
        {block.type === 'editor' ? (
          <>
            <TextField
              label="Título"
              value={editor.title}
              onChange={(_, v) => setEditor((e) => ({ ...e, title: v ?? '' }))}
            />
            <TextField
              label="Placeholder"
              value={editor.placeholder}
              onChange={(_, v) => setEditor((e) => ({ ...e, placeholder: v ?? '' }))}
            />
            <TextField
              label="Altura mínima (px)"
              type="number"
              value={String(editor.minHeightPx)}
              onChange={(_, v) => {
                const n = parseInt(String(v ?? ''), 10);
                if (isNaN(n)) return;
                setEditor((e) => ({ ...e, minHeightPx: Math.min(2000, Math.max(40, n)) }));
              }}
            />
            <Toggle
              label="Somente leitura (na página)"
              checked={editor.readOnly}
              onChange={(_, c) => setEditor((e) => ({ ...e, readOnly: Boolean(c) }))}
            />
            <Text variant="smallPlus" styles={{ root: { fontWeight: 600, marginTop: 4 } }}>
              Permitir na barra do editor e na exibição
            </Text>
            <Toggle
              label="Imagens"
              checked={editor.allowImages}
              onChange={(_, c) => setEditor((e) => ({ ...e, allowImages: Boolean(c) }))}
            />
            <Toggle
              label="Links"
              checked={editor.allowLinks}
              onChange={(_, c) => setEditor((e) => ({ ...e, allowLinks: Boolean(c) }))}
            />
            <Toggle
              label="Tabelas (HTML; use modo HTML abaixo)"
              checked={editor.allowTables}
              onChange={(_, c) => setEditor((e) => ({ ...e, allowTables: Boolean(c) }))}
            />
            <Toggle
              label="Listas"
              checked={editor.allowLists}
              onChange={(_, c) => setEditor((e) => ({ ...e, allowLists: Boolean(c) }))}
            />
            <Toggle
              label="Cabeçalhos"
              checked={editor.allowHeaders}
              onChange={(_, c) => setEditor((e) => ({ ...e, allowHeaders: Boolean(c) }))}
            />
            <Toggle
              label="Vídeo / embed (YouTube, Vimeo)"
              checked={editor.allowVideoEmbed}
              onChange={(_, c) => setEditor((e) => ({ ...e, allowVideoEmbed: Boolean(c) }))}
            />
            <Toggle
              label="Editar como HTML (tabelas e markup avançado)"
              checked={editorHtmlSourceMode}
              onChange={(_, c) => setEditorHtmlSourceMode(Boolean(c))}
            />
            {editorHtmlSourceMode ? (
              <TextField
                label="HTML"
                multiline
                rows={14}
                value={editor.html}
                onChange={(_, v) => setEditor((e) => ({ ...e, html: v ?? '' }))}
                description="Na exibição o HTML é filtrado pelas permissões acima."
              />
            ) : (
              <ListPageRichQuillEditor
                value={editor.html}
                onChange={(html) => setEditor((e) => ({ ...e, html }))}
                placeholder={editor.placeholder}
                permissions={{
                  allowImages: editor.allowImages,
                  allowLinks: editor.allowLinks,
                  allowLists: editor.allowLists,
                  allowHeaders: editor.allowHeaders,
                  allowVideoEmbed: editor.allowVideoEmbed,
                }}
              />
            )}
            <Text variant="small" styles={{ root: { color: '#605e5c', lineHeight: 1.45 } }}>
              Editor: React Quill. O conteúdo é guardado em HTML e sanitizado na página pública conforme as
              opções permitidas.
            </Text>
          </>
        ) : null}
        {block.type === 'sectionTitle' ? (
          <>
            <TextField
              label="Título"
              value={sectionTitle.title}
              onChange={(_, v) => setSectionTitle((s) => ({ ...s, title: v ?? '' }))}
            />
            <TextField
              label="Subtítulo"
              multiline
              rows={2}
              value={sectionTitle.subtitle}
              onChange={(_, v) => setSectionTitle((s) => ({ ...s, subtitle: v ?? '' }))}
            />
            <TextField
              label="Ícone"
              value={sectionTitle.iconName}
              onChange={(_, v) => setSectionTitle((s) => ({ ...s, iconName: v ?? '' }))}
              description="Nome do ícone Fluent UI (ex.: Page, Info)."
            />
            <Dropdown
              label="Alinhamento"
              selectedKey={sectionTitle.align}
              options={ALIGN_OPTIONS}
              onChange={(_, opt) => {
                const k = opt?.key as TListPageBannerContentAlign | undefined;
                if (k === 'left' || k === 'center' || k === 'right') {
                  setSectionTitle((s) => ({ ...s, align: k }));
                }
              }}
            />
            <Toggle
              label="Exibir linha divisória"
              checked={sectionTitle.showDivider}
              onChange={(_, c) => setSectionTitle((s) => ({ ...s, showDivider: Boolean(c) }))}
            />
            <Dropdown
              label="Tamanho"
              selectedKey={sectionTitle.size}
              options={SECTION_TITLE_SIZE_OPTIONS}
              onChange={(_, opt) => {
                const k = opt?.key as TListPageSectionTitleSize | undefined;
                if (k === 'sm' || k === 'md' || k === 'lg') {
                  setSectionTitle((s) => ({ ...s, size: k }));
                }
              }}
            />
            <TextField
              label="Margem superior (px)"
              type="number"
              value={String(sectionTitle.marginTopPx)}
              onChange={(_, v) => {
                const n = parseInt(String(v ?? ''), 10);
                if (isNaN(n)) return;
                setSectionTitle((s) => ({ ...s, marginTopPx: Math.min(120, Math.max(0, n)) }));
              }}
            />
            <TextField
              label="Margem inferior (px)"
              type="number"
              value={String(sectionTitle.marginBottomPx)}
              onChange={(_, v) => {
                const n = parseInt(String(v ?? ''), 10);
                if (isNaN(n)) return;
                setSectionTitle((s) => ({ ...s, marginBottomPx: Math.min(120, Math.max(0, n)) }));
              }}
            />
          </>
        ) : null}
        {block.type === 'alert' ? (
          <>
            <TextField
              label="Título"
              value={alertCfg.title}
              onChange={(_, v) => setAlertCfg((a) => ({ ...a, title: v ?? '' }))}
            />
            <TextField
              label="Mensagem"
              multiline
              rows={4}
              value={alertCfg.message}
              onChange={(_, v) => setAlertCfg((a) => ({ ...a, message: v ?? '' }))}
            />
            <Dropdown
              label="Tipo"
              selectedKey={alertCfg.variant}
              options={ALERT_VARIANT_OPTIONS}
              onChange={(_, opt) => {
                const k = opt?.key as TListPageAlertVariant | undefined;
                if (k === 'info' || k === 'success' || k === 'warning' || k === 'error') {
                  setAlertCfg((a) => ({ ...a, variant: k }));
                }
              }}
            />
            <TextField
              label="Ícone"
              value={alertCfg.iconName}
              onChange={(_, v) => setAlertCfg((a) => ({ ...a, iconName: v ?? '' }))}
              description="Opcional. Vazio = ícone padrão do tipo."
            />
            <Toggle
              label="Pode ser fechado"
              checked={alertCfg.dismissible}
              onChange={(_, c) => setAlertCfg((a) => ({ ...a, dismissible: Boolean(c) }))}
            />
            <Toggle
              label="Destaque (borda e fundo)"
              checked={alertCfg.emphasized}
              onChange={(_, c) => setAlertCfg((a) => ({ ...a, emphasized: Boolean(c) }))}
            />
            <TextField
              label="URL do link (opcional)"
              value={alertCfg.linkUrl}
              onChange={(_, v) => setAlertCfg((a) => ({ ...a, linkUrl: v ?? '' }))}
            />
            <TextField
              label="Texto do link"
              value={alertCfg.linkText}
              onChange={(_, v) => setAlertCfg((a) => ({ ...a, linkText: v ?? '' }))}
            />
          </>
        ) : null}
      </Stack>
    </Panel>
  );
};
