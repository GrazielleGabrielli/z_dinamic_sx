import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
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
  IconButton,
  Separator,
} from '@fluentui/react';
import type {
  IListPageAlertBlockConfig,
  IListPageAlertCountRule,
  IListPageBannerBlockConfig,
  IListPageBlock,
  IListPageButtonItemConfig,
  IListPageButtonsBlockConfig,
  IListPageRichEditorBlockConfig,
  IListPageSectionTitleBlockConfig,
  TListPageAlertCountOp,
  TListPageAlertVariant,
  TListPageBannerContentAlign,
  TListPageButtonActionKind,
  TListPageSectionTitleSize,
} from '../../core/config/types';
import {
  defaultAlertConfig,
  defaultBannerConfig,
  defaultButtonsConfig,
  defaultRichEditorConfig,
  defaultSectionTitleConfig,
  MAX_ALERT_COUNT_RULES,
  MAX_LIST_PAGE_BUTTONS,
  sanitizeButtonsConfig,
} from '../../core/listPage/listPageBlockConfigUtils';
import { ListPageRichQuillEditor } from './ListPageRichQuillEditor';
import { AlertCountRuleFilterFields } from './AlertCountRuleFilterFields';
import { FieldsService, type IFieldMetadata } from '../../../../services';
import { mergeCountRuleODataFromStructured } from '../../core/listPage/alertCountRuleFilterOData';

export interface IListPageBlockConfigPanelProps {
  isOpen: boolean;
  block: IListPageBlock | null;
  /** Lista da vista (contagem OData nas regras). */
  listTitle?: string;
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

const ALERT_COUNT_OP_OPTIONS: IDropdownOption[] = [
  { key: 'eq', text: 'Igual a (=)' },
  { key: 'ne', text: 'Diferente de (≠)' },
  { key: 'gt', text: 'Maior que (>)' },
  { key: 'ge', text: 'Maior ou igual (≥)' },
  { key: 'lt', text: 'Menor que (<)' },
  { key: 'le', text: 'Menor ou igual (≤)' },
];

const INHERIT_TRISTATE_OPTIONS: IDropdownOption[] = [
  { key: 'inherit', text: 'Herdar do padrão' },
  { key: 'yes', text: 'Sim' },
  { key: 'no', text: 'Não' },
];

function newAlertCountRule(): IListPageAlertCountRule {
  return {
    id: `acr_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`,
    odataFilter: '',
    countOp: 'eq',
    count: 0,
  };
}

function newListPageButtonDraft(): IListPageButtonItemConfig {
  return {
    id: `lpbtn_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`,
    label: '',
    actionKind: 'redirect',
    url: '',
    openInNewTab: false,
  };
}

const BUTTON_ACTION_OPTIONS: IDropdownOption[] = [
  { key: 'redirect', text: 'Redirecionar (URL)' },
  { key: 'reload', text: 'Recarregar página' },
];

function panelHeaderForBlockType(t: IListPageBlock['type']): string {
  if (t === 'banner') return 'Banner';
  if (t === 'editor') return 'Editor de conteúdo';
  if (t === 'sectionTitle') return 'Título de seção';
  if (t === 'alert') return 'Alerta / aviso';
  if (t === 'buttons') return 'Botões';
  return '';
}

function panelWidthForBlockType(t: IListPageBlock['type']): string {
  if (t === 'editor') return '620px';
  if (t === 'alert') return '560px';
  if (t === 'buttons') return '520px';
  return '480px';
}

export const ListPageBlockConfigPanel: React.FC<IListPageBlockConfigPanelProps> = ({
  isOpen,
  block,
  listTitle = '',
  onDismiss,
  onApply,
}) => {
  const [banner, setBanner] = useState<IListPageBannerBlockConfig>(defaultBannerConfig);
  const [editor, setEditor] = useState<IListPageRichEditorBlockConfig>(defaultRichEditorConfig);
  const [editorHtmlSourceMode, setEditorHtmlSourceMode] = useState(false);
  const [sectionTitle, setSectionTitle] = useState<IListPageSectionTitleBlockConfig>(defaultSectionTitleConfig);
  const [alertCfg, setAlertCfg] = useState<IListPageAlertBlockConfig>(defaultAlertConfig);
  const [alertListFields, setAlertListFields] = useState<IFieldMetadata[] | undefined>(undefined);
  const [buttonsCfg, setButtonsCfg] = useState<IListPageButtonsBlockConfig>(() => defaultButtonsConfig());
  const fieldsService = useMemo(() => new FieldsService(), []);

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
      const src = block.alert ? { ...block.alert } : defaultAlertConfig();
      src.countRules = block.alert?.countRules?.length
        ? block.alert.countRules.map((r) => ({ ...r }))
        : [];
      setAlertCfg(src);
    }
    if (block.type === 'buttons') {
      const items = block.buttons?.items?.length
        ? block.buttons.items.map((it) => ({ ...it }))
        : defaultButtonsConfig().items.map((it) => ({ ...it }));
      setButtonsCfg({ items });
    }
  }, [isOpen, block]);

  useEffect(() => {
    if (!isOpen || block?.type !== 'alert' || !listTitle.trim()) {
      setAlertListFields(undefined);
      return;
    }
    fieldsService
      .getVisibleFields(listTitle.trim())
      .then((f) => setAlertListFields(f))
      .catch(() => setAlertListFields([]));
  }, [isOpen, block?.type, listTitle, fieldsService]);

  if (
    !block ||
    (block.type !== 'banner' &&
      block.type !== 'editor' &&
      block.type !== 'sectionTitle' &&
      block.type !== 'alert' &&
      block.type !== 'buttons')
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
    } else if (block.type === 'buttons') {
      onApply({ ...block, buttons: sanitizeButtonsConfig({ items: buttonsCfg.items }) });
    } else {
      const by = new Map((alertListFields ?? []).map((f) => [f.InternalName, f]));
      const rules = (alertCfg.countRules ?? []).map((r) => mergeCountRuleODataFromStructured(r, by));
      const alert: IListPageAlertBlockConfig = {
        ...alertCfg,
        countRules: rules.length ? rules : undefined,
      };
      onApply({ ...block, alert });
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
            {!listTitle.trim() ? (
              <Text variant="small" styles={{ root: { color: '#a4262c', lineHeight: 1.45 } }}>
                Defina o título da lista na fonte de dados da vista para as regras de contagem funcionarem.
              </Text>
            ) : (
              <Text variant="small" styles={{ root: { color: '#605e5c', lineHeight: 1.45 } }}>
                Contagem na lista: <strong>{listTitle.trim()}</strong>
              </Text>
            )}
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
            <Separator />
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                Regras por contagem (ordem importa)
              </Text>
              <IconButton
                iconProps={{ iconName: 'Add' }}
                title="Adicionar regra"
                ariaLabel="Adicionar regra"
                disabled={(alertCfg.countRules?.length ?? 0) >= MAX_ALERT_COUNT_RULES}
                onClick={() =>
                  setAlertCfg((a) => ({
                    ...a,
                    countRules: [...(a.countRules ?? []), newAlertCountRule()],
                  }))
                }
              />
            </Stack>
            <Stack
              tokens={{ childrenGap: 8 }}
              styles={{
                root: {
                  maxWidth: '100%',
                  marginTop: 2,
                },
              }}
            >
             
            </Stack>
            {(alertCfg.countRules ?? []).map((rule, idx) => (
              <Stack
                key={rule.id}
                tokens={{ childrenGap: 8 }}
                styles={{
                  root: {
                    border: '1px solid #edebe9',
                    borderRadius: 4,
                    padding: 10,
                    background: '#faf9f8',
                  },
                }}
              >
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                  <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                    Regra {idx + 1}
                  </Text>
                  <Stack horizontal verticalAlign="center">
                    <IconButton
                      iconProps={{ iconName: 'ChevronUpSmall' }}
                      title="Subir"
                      ariaLabel="Subir regra"
                      disabled={idx === 0}
                      onClick={() =>
                        setAlertCfg((a) => {
                          const rules = [...(a.countRules ?? [])];
                          if (idx <= 0) return a;
                          [rules[idx - 1], rules[idx]] = [rules[idx], rules[idx - 1]];
                          return { ...a, countRules: rules };
                        })
                      }
                    />
                    <IconButton
                      iconProps={{ iconName: 'ChevronDownSmall' }}
                      title="Descer"
                      ariaLabel="Descer regra"
                      disabled={idx >= (alertCfg.countRules?.length ?? 0) - 1}
                      onClick={() =>
                        setAlertCfg((a) => {
                          const rules = [...(a.countRules ?? [])];
                          if (idx >= rules.length - 1) return a;
                          [rules[idx + 1], rules[idx]] = [rules[idx], rules[idx + 1]];
                          return { ...a, countRules: rules };
                        })
                      }
                    />
                    <IconButton
                      iconProps={{ iconName: 'Delete' }}
                      title="Remover"
                      ariaLabel="Remover regra"
                      onClick={() =>
                        setAlertCfg((a) => ({
                          ...a,
                          countRules: (a.countRules ?? []).filter((_, i) => i !== idx),
                        }))
                      }
                    />
                  </Stack>
                </Stack>
                <AlertCountRuleFilterFields
                  listTitle={listTitle}
                  fields={alertListFields}
                  rule={rule}
                  onRuleChange={(next) => {
                    setAlertCfg((a) => {
                      const rules = [...(a.countRules ?? [])];
                      rules[idx] = next;
                      return { ...a, countRules: rules };
                    });
                  }}
                />
                <Text variant="small" styles={{ root: { color: '#605e5c', lineHeight: 1.45 } }}>
                  Sem campo selecionado = contar todos os itens (até 5000 por pedido ao SharePoint).
                </Text>
                <Stack horizontal tokens={{ childrenGap: 10 }}>
                  <Stack.Item grow={1} styles={{ root: { minWidth: 200 } }}>
                    <Dropdown
                      label="Condição"
                      selectedKey={rule.countOp}
                      options={ALERT_COUNT_OP_OPTIONS}
                      onChange={(_, opt) => {
                        const k = opt?.key as TListPageAlertCountOp | undefined;
                        if (!k) return;
                        setAlertCfg((a) => {
                          const rules = [...(a.countRules ?? [])];
                          rules[idx] = { ...rules[idx], countOp: k };
                          return { ...a, countRules: rules };
                        });
                      }}
                    />
                  </Stack.Item>
                  <Stack.Item grow={1}>
                    <TextField
                      label="Valor"
                      type="number"
                      value={String(rule.count)}
                      onChange={(_, v) => {
                        const n = parseInt(String(v ?? ''), 10);
                        if (isNaN(n)) return;
                        setAlertCfg((a) => {
                          const rules = [...(a.countRules ?? [])];
                          rules[idx] = { ...rules[idx], count: Math.min(5000, Math.max(0, n)) };
                          return { ...a, countRules: rules };
                        });
                      }}
                    />
                  </Stack.Item>
                </Stack>
                <Text variant="small" styles={{ root: { fontWeight: 600, marginTop: 4 } }}>
                  Aspeto quando esta regra coincidir (opcional; vazio mantém o padrão)
                </Text>
                <TextField
                  label="Título"
                  value={rule.title ?? ''}
                  onChange={(_, v) => {
                    const s = v ?? '';
                    setAlertCfg((a) => {
                      const rules = [...(a.countRules ?? [])];
                      const next = { ...rules[idx] };
                      if (s.trim()) next.title = s;
                      else delete next.title;
                      rules[idx] = next;
                      return { ...a, countRules: rules };
                    });
                  }}
                />
                <TextField
                  label="Mensagem"
                  multiline
                  rows={2}
                  value={rule.message ?? ''}
                  onChange={(_, v) => {
                    const s = v ?? '';
                    setAlertCfg((a) => {
                      const rules = [...(a.countRules ?? [])];
                      const next = { ...rules[idx] };
                      if (s.trim()) next.message = s;
                      else delete next.message;
                      rules[idx] = next;
                      return { ...a, countRules: rules };
                    });
                  }}
                />
                <Dropdown
                  label="Tipo"
                  selectedKey={rule.variant ?? 'inherit'}
                  options={[{ key: 'inherit', text: '(padrão acima)' }, ...ALERT_VARIANT_OPTIONS]}
                  onChange={(_, opt) => {
                    const k = opt?.key as string | undefined;
                    setAlertCfg((a) => {
                      const rules = [...(a.countRules ?? [])];
                      const next = { ...rules[idx] };
                      if (k === 'inherit' || !k) delete next.variant;
                      else if (k === 'info' || k === 'success' || k === 'warning' || k === 'error') {
                        next.variant = k;
                      }
                      rules[idx] = next;
                      return { ...a, countRules: rules };
                    });
                  }}
                />
                <TextField
                  label="Ícone"
                  value={rule.iconName ?? ''}
                  onChange={(_, v) => {
                    const s = (v ?? '').trim();
                    setAlertCfg((a) => {
                      const rules = [...(a.countRules ?? [])];
                      const next = { ...rules[idx] };
                      if (s) next.iconName = s;
                      else delete next.iconName;
                      rules[idx] = next;
                      return { ...a, countRules: rules };
                    });
                  }}
                />
                <Dropdown
                  label="Pode ser fechado"
                  selectedKey={
                    rule.dismissible === undefined ? 'inherit' : rule.dismissible ? 'yes' : 'no'
                  }
                  options={INHERIT_TRISTATE_OPTIONS}
                  onChange={(_, opt) => {
                    const k = String(opt?.key ?? 'inherit');
                    setAlertCfg((a) => {
                      const rules = [...(a.countRules ?? [])];
                      const next = { ...rules[idx] };
                      if (k === 'inherit') delete next.dismissible;
                      else next.dismissible = k === 'yes';
                      rules[idx] = next;
                      return { ...a, countRules: rules };
                    });
                  }}
                />
                <Dropdown
                  label="Destaque (borda)"
                  selectedKey={rule.emphasized === undefined ? 'inherit' : rule.emphasized ? 'yes' : 'no'}
                  options={INHERIT_TRISTATE_OPTIONS}
                  onChange={(_, opt) => {
                    const k = String(opt?.key ?? 'inherit');
                    setAlertCfg((a) => {
                      const rules = [...(a.countRules ?? [])];
                      const next = { ...rules[idx] };
                      if (k === 'inherit') delete next.emphasized;
                      else next.emphasized = k === 'yes';
                      rules[idx] = next;
                      return { ...a, countRules: rules };
                    });
                  }}
                />
                <TextField
                  label="URL do link"
                  value={rule.linkUrl ?? ''}
                  onChange={(_, v) => {
                    const s = (v ?? '').trim();
                    setAlertCfg((a) => {
                      const rules = [...(a.countRules ?? [])];
                      const next = { ...rules[idx] };
                      if (s) next.linkUrl = s;
                      else delete next.linkUrl;
                      rules[idx] = next;
                      return { ...a, countRules: rules };
                    });
                  }}
                />
                <TextField
                  label="Texto do link"
                  value={rule.linkText ?? ''}
                  onChange={(_, v) => {
                    const s = v ?? '';
                    setAlertCfg((a) => {
                      const rules = [...(a.countRules ?? [])];
                      const next = { ...rules[idx] };
                      if (s.trim()) next.linkText = s;
                      else delete next.linkText;
                      rules[idx] = next;
                      return { ...a, countRules: rules };
                    });
                  }}
                />
              </Stack>
            ))}
          </>
        ) : null}
        {block.type === 'buttons' ? (
          <>
            <Text variant="small" styles={{ root: { color: '#605e5c', lineHeight: 1.45 } }}>
              Tipos: abrir uma URL (mesma aba ou nova) ou recarregar a página atual. Ao aplicar, linhas sem texto ou
              redirecionamento sem URL válida são ignoradas (máximo {MAX_LIST_PAGE_BUTTONS} botões).
            </Text>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                Botões
              </Text>
              <IconButton
                iconProps={{ iconName: 'Add' }}
                title="Adicionar botão"
                ariaLabel="Adicionar botão"
                disabled={buttonsCfg.items.length >= MAX_LIST_PAGE_BUTTONS}
                onClick={() =>
                  setButtonsCfg((b) => ({ ...b, items: [...b.items, newListPageButtonDraft()] }))
                }
              />
            </Stack>
            {buttonsCfg.items.map((item, idx) => (
              <Stack
                key={item.id}
                tokens={{ childrenGap: 8 }}
                styles={{
                  root: {
                    border: '1px solid #edebe9',
                    borderRadius: 4,
                    padding: 10,
                    background: '#faf9f8',
                  },
                }}
              >
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                  <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                    Botão {idx + 1}
                  </Text>
                  <Stack horizontal verticalAlign="center">
                    <IconButton
                      iconProps={{ iconName: 'ChevronUpSmall' }}
                      title="Subir"
                      ariaLabel="Subir botão"
                      disabled={idx === 0}
                      onClick={() =>
                        setButtonsCfg((b) => {
                          const items = [...b.items];
                          if (idx <= 0) return b;
                          [items[idx - 1], items[idx]] = [items[idx], items[idx - 1]];
                          return { ...b, items };
                        })
                      }
                    />
                    <IconButton
                      iconProps={{ iconName: 'ChevronDownSmall' }}
                      title="Descer"
                      ariaLabel="Descer botão"
                      disabled={idx >= buttonsCfg.items.length - 1}
                      onClick={() =>
                        setButtonsCfg((b) => {
                          const items = [...b.items];
                          if (idx >= items.length - 1) return b;
                          [items[idx + 1], items[idx]] = [items[idx], items[idx + 1]];
                          return { ...b, items };
                        })
                      }
                    />
                    <IconButton
                      iconProps={{ iconName: 'Delete' }}
                      title="Remover"
                      ariaLabel="Remover botão"
                      disabled={buttonsCfg.items.length <= 1}
                      onClick={() =>
                        setButtonsCfg((b) => ({
                          ...b,
                          items: b.items.filter((_, i) => i !== idx),
                        }))
                      }
                    />
                  </Stack>
                </Stack>
                <TextField
                  label="Texto do botão"
                  value={item.label}
                  onChange={(_, v) => {
                    const s = v ?? '';
                    setButtonsCfg((b) => {
                      const items = [...b.items];
                      items[idx] = { ...items[idx], label: s };
                      return { ...b, items };
                    });
                  }}
                />
                <Dropdown
                  label="Tipo"
                  selectedKey={item.actionKind}
                  options={BUTTON_ACTION_OPTIONS}
                  onChange={(_, opt) => {
                    const k = opt?.key as TListPageButtonActionKind | undefined;
                    if (k !== 'redirect' && k !== 'reload') return;
                    setButtonsCfg((b) => {
                      const items = [...b.items];
                      const cur = { ...items[idx], actionKind: k };
                      if (k === 'reload') {
                        delete cur.url;
                        delete cur.openInNewTab;
                      } else {
                        cur.url = cur.url ?? '';
                        cur.openInNewTab = Boolean(cur.openInNewTab);
                      }
                      items[idx] = cur as IListPageButtonItemConfig;
                      return { ...b, items };
                    });
                  }}
                />
                {item.actionKind === 'redirect' ? (
                  <>
                    <TextField
                      label="URL de destino"
                      value={item.url ?? ''}
                      onChange={(_, v) => {
                        const s = v ?? '';
                        setButtonsCfg((b) => {
                          const items = [...b.items];
                          items[idx] = { ...items[idx], url: s };
                          return { ...b, items };
                        });
                      }}
                    />
                    <Toggle
                      label="Abrir em nova aba"
                      checked={item.openInNewTab === true}
                      onChange={(_, c) =>
                        setButtonsCfg((b) => {
                          const items = [...b.items];
                          items[idx] = { ...items[idx], openInNewTab: Boolean(c) };
                          return { ...b, items };
                        })
                      }
                    />
                  </>
                ) : null}
              </Stack>
            ))}
          </>
        ) : null}
      </Stack>
    </Panel>
  );
};


