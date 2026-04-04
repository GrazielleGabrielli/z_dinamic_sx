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
  IListPageBannerBlockConfig,
  IListPageBlock,
  IListPageRichEditorBlockConfig,
  TListPageBannerContentAlign,
} from '../../core/config/types';
import {
  defaultBannerConfig,
  defaultRichEditorConfig,
} from '../../core/listPage/listPageBlockConfigUtils';

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

export const ListPageBlockConfigPanel: React.FC<IListPageBlockConfigPanelProps> = ({
  isOpen,
  block,
  onDismiss,
  onApply,
}) => {
  const [banner, setBanner] = useState<IListPageBannerBlockConfig>(defaultBannerConfig);
  const [editor, setEditor] = useState<IListPageRichEditorBlockConfig>(defaultRichEditorConfig);

  useEffect(() => {
    if (!isOpen || !block) return;
    if (block.type === 'banner') {
      setBanner(block.banner ? { ...block.banner } : defaultBannerConfig());
    }
    if (block.type === 'editor') {
      setEditor(block.editor ? { ...block.editor } : defaultRichEditorConfig());
    }
  }, [isOpen, block]);

  if (!block || (block.type !== 'banner' && block.type !== 'editor')) {
    return null;
  }

  const handleSave = (): void => {
    if (block.type === 'banner') {
      onApply({ ...block, banner });
    } else {
      onApply({ ...block, editor });
    }
  };

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.custom}
      customWidth="480px"
      headerText={block.type === 'banner' ? 'Banner' : 'Editor de conteúdo'}
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
        ) : (
          <>
            <TextField
              label="Título"
              value={editor.title}
              onChange={(_, v) => setEditor((e) => ({ ...e, title: v ?? '' }))}
            />
            <TextField
              label="Conteúdo rico (HTML)"
              multiline
              rows={14}
              value={editor.html}
              onChange={(_, v) => setEditor((e) => ({ ...e, html: v ?? '' }))}
              description="O HTML é filtrado na página conforme as permissões abaixo."
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
              label="Somente leitura"
              checked={editor.readOnly}
              onChange={(_, c) => setEditor((e) => ({ ...e, readOnly: Boolean(c) }))}
            />
            <Text variant="smallPlus" styles={{ root: { fontWeight: 600, marginTop: 4 } }}>
              Permitir na exibição
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
              label="Tabelas"
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
          </>
        )}
      </Stack>
    </Panel>
  );
};
