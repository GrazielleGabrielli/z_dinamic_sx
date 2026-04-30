import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  Shimmer,
  ShimmerElementType,
  ProgressIndicator,
  IconButton,
  useTheme,
  Toggle,
  ChoiceGroup,
  type IChoiceGroupOption,
  TextField,
  Checkbox,
  DefaultButton,
  MessageBar,
  MessageBarType,
} from '@fluentui/react';
import { Dropdown } from '@fluentui/react';
import { filterSiteGroupsByNameQuery, type IGroupDetails } from '../../../../services';
import type {
  TFormStepLayoutKind,
  TFormStepNavButtonsKind,
  TFormDataLoadingUiKind,
  TFormSubmitLoadingUiKind,
  TFormAttachmentUploadLayoutKind,
  TFormAttachmentFilePreviewKind,
  TFormHistoryLayoutKind,
  TFormHistoryPresentationKind,
  TFormHistoryButtonKind,
  TFormCustomButtonPaletteSlot,
} from '../../core/config/types/formManager';
import { resolveStepUiAccentColor } from '../../core/formManager/formCustomButtonTheme';
import { FormStepLayoutPicker, FormStepNavButtonsPicker } from './FormStepLayoutUi';
import { ThemePaletteSlotDropdown } from './ThemePaletteSlotDropdown';
import {
  FormDataLoadingView,
  FORM_DATA_LOADING_DROPDOWN_OPTIONS,
  FORM_SUBMIT_LOADING_DROPDOWN_OPTIONS,
} from './FormLoadingUi';

function normSpGroupTitle(s: string): string {
  return s.trim().toLowerCase();
}

const HISTORY_BUTTON_KIND_OPTIONS: IChoiceGroupOption[] = [
  { key: 'text', text: 'Só texto' },
  { key: 'icon', text: 'Só ícone' },
  { key: 'iconAndText', text: 'Ícone e texto' },
];

export const FORM_ATTACHMENT_LAYOUT_DROPDOWN_OPTIONS: {
  key: TFormAttachmentUploadLayoutKind;
  text: string;
}[] = [
  { key: 'default', text: 'Clássico (input nativo)' },
  { key: 'dropzone', text: 'Zona destacada (largar / clicar)' },
  { key: 'card', text: 'Cartão com ícone e sombra' },
  { key: 'ribbon', text: 'Faixa azul + área de largar' },
  { key: 'compact', text: 'Compacto (botão + chips)' },
];

export interface IFormAttachmentExtensionPreset {
  ext: string;
  label: string;
}

export interface IFormAttachmentExtensionGroup {
  title: string;
  hint?: string;
  items: IFormAttachmentExtensionPreset[];
}

export const FORM_ATTACHMENT_EXTENSION_GROUPS: IFormAttachmentExtensionGroup[] = [
  {
    title: 'PDF e documentos Word',
    items: [
      { ext: 'pdf', label: 'PDF' },
      { ext: 'doc', label: 'Word .doc' },
      { ext: 'docx', label: 'Word .docx' },
    ],
  },
  {
    title: 'Excel',
    items: [
      { ext: 'xls', label: '.xls' },
      { ext: 'xlsx', label: '.xlsx' },
    ],
  },
  {
    title: 'PowerPoint',
    items: [
      { ext: 'ppt', label: '.ppt' },
      { ext: 'pptx', label: '.pptx' },
    ],
  },
  {
    title: 'Imagens',
    hint: 'Raster e vetorial',
    items: [
      { ext: 'png', label: 'PNG' },
      { ext: 'jpg', label: 'JPEG .jpg' },
      { ext: 'jpeg', label: 'JPEG .jpeg' },
      { ext: 'gif', label: 'GIF' },
      { ext: 'webp', label: 'WebP' },
      { ext: 'svg', label: 'SVG' },
    ],
  },
  {
    title: 'Texto e tabelas',
    items: [
      { ext: 'txt', label: 'Texto .txt' },
      { ext: 'csv', label: 'CSV' },
    ],
  },
  {
    title: 'Arquivos e correio',
    items: [
      { ext: 'zip', label: 'ZIP' },
      { ext: 'msg', label: 'Outlook .msg' },
    ],
  },
  {
    title: 'Vídeo',
    items: [{ ext: 'mp4', label: 'MP4' }],
  },
];

export const FORM_ATTACHMENT_FILE_PREVIEW_DROPDOWN_OPTIONS: {
  key: TFormAttachmentFilePreviewKind;
  text: string;
}[] = [
  { key: 'nameOnly', text: 'Só nome do ficheiro' },
  { key: 'nameAndSize', text: 'Nome e tamanho (padrão)' },
  { key: 'iconAndName', text: 'Ícone por tipo + nome (+ tamanho)' },
  { key: 'thumbnailAndName', text: 'Miniatura (imagem) ou ícone + nome' },
  { key: 'thumbnailLarge', text: 'Pré-visualização grande (cartão por ficheiro)' },
];

export const FORM_HISTORY_LAYOUT_DROPDOWN_OPTIONS: {
  key: TFormHistoryLayoutKind;
  text: string;
}[] = [
  { key: 'list', text: 'Lista (blocos empilhados)' },
  { key: 'timeline', text: 'Linha do tempo (eixo vertical)' },
  { key: 'cards', text: 'Cartões (sombra e destaque)' },
  { key: 'compact', text: 'Compacto (denso, uma linha por meta)' },
];

function HistoryLayoutPreview({ kind }: { kind: TFormHistoryLayoutKind }): JSX.Element {
  const samples = [
    { title: 'Registo 1', meta: '01/04/2026 10:02 · Ana', hint: 'Texto do campo multilinhas…' },
    { title: 'Registo 2', meta: '01/04/2026 09:10 · Bruno', hint: 'Outro registo…' },
  ];
  const wrap = (child: React.ReactNode): JSX.Element => (
    <Stack
      styles={{
        root: {
          border: '1px solid #edebe9',
          borderRadius: 4,
          padding: 12,
          background: '#faf9f8',
          marginTop: 8,
        },
      }}
    >
      <Text variant="small" styles={{ root: { fontWeight: 600, color: '#605e5c', marginBottom: 8 } }}>
        Pré-visualização do estilo
      </Text>
      {child}
    </Stack>
  );
  if (kind === 'timeline') {
    return wrap(
      <div style={{ position: 'relative', paddingLeft: 22 }}>
        <div
          style={{
            position: 'absolute',
            left: 5,
            top: 6,
            bottom: 6,
            width: 2,
            background: '#e1dfdd',
          }}
        />
        <Stack tokens={{ childrenGap: 14 }}>
          {samples.map((s, i) => (
            <div key={i} style={{ position: 'relative' }}>
              <div
                style={{
                  position: 'absolute',
                  left: -19,
                  top: 2,
                  width: 12,
                  height: 12,
                  borderRadius: '50%',
                  background: '#0078d4',
                  border: '2px solid #fff',
                  boxShadow: '0 0 0 1px #c8c6c4',
                }}
              />
              <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                {s.title}
              </Text>
              <Text variant="small" styles={{ root: { color: '#605e5c', fontSize: 11 } }}>
                {s.meta}
              </Text>
              <Text variant="small" styles={{ root: { color: '#a19f9d', fontSize: 11 } }}>
                {s.hint}
              </Text>
            </div>
          ))}
        </Stack>
      </div>
    );
  }
  if (kind === 'cards') {
    return wrap(
      <Stack tokens={{ childrenGap: 10 }}>
        {samples.map((s, i) => (
          <div
            key={i}
            style={{
              padding: 14,
              borderRadius: 8,
              background: '#fff',
              boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
              border: '1px solid #edebe9',
            }}
          >
            <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
              {s.title}
            </Text>
            <Text variant="small" styles={{ root: { color: '#605e5c', fontSize: 11, marginTop: 4 } }}>
              {s.meta}
            </Text>
          </div>
        ))}
      </Stack>
    );
  }
  if (kind === 'compact') {
    return wrap(
      <Stack tokens={{ childrenGap: 0 }}>
        {samples.map((s, i) => (
          <div
            key={i}
            style={{
              padding: '6px 0',
              borderBottom: i < samples.length - 1 ? '1px solid #edebe9' : undefined,
            }}
          >
            <Text variant="small" styles={{ root: { fontSize: 11, color: '#323130' } }}>
              <span style={{ fontWeight: 600 }}>{s.title}</span>
              <span style={{ color: '#605e5c' }}> · {s.meta}</span>
            </Text>
          </div>
        ))}
      </Stack>
    );
  }
  return wrap(
    <Stack tokens={{ childrenGap: 8 }}>
      {samples.map((s, i) => (
        <div
          key={i}
          style={{
            padding: '10px 12px',
            borderRadius: 4,
            border: '1px solid #edebe9',
            background: '#ffffff',
          }}
        >
          <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
            {s.title}
          </Text>
          <Text variant="small" styles={{ root: { color: '#605e5c', fontSize: 11, marginTop: 4 } }}>
            {s.meta}
          </Text>
          <Text variant="small" styles={{ root: { color: '#a19f9d', fontSize: 11, marginTop: 4 } }}>
            {s.hint}
          </Text>
        </div>
      ))}
    </Stack>
  );
}

const loadingCardStyles = (): { root: Record<string, string | number> } => ({
  root: {
    padding: 16,
    borderRadius: 4,
    border: '1px solid #edebe9',
    background: '#faf9f8',
  },
});

const SECTION_IDS = {
  loadData: 'loadData',
  submitLoading: 'submitLoading',
  steps: 'steps',
  historyAudit: 'historyAudit',
} as const;

/** Mesmo padrão visual do painel «Layout da página (modo lista)». */
export function FormManagerCollapseSection(props: {
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

export function FormManagerComponentsLoadingLayouts(): JSX.Element {
  const [idx, setIdx] = useState(0);
  useEffect(() => {
    const id = window.setInterval(() => setIdx((i) => (i + 1) % 3), 2600);
    return () => window.clearInterval(id);
  }, []);

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        A carregar metadados dos campos da lista…
      </Text>
      {idx === 0 && (
        <Stack horizontalAlign="center" tokens={{ childrenGap: 16 }} styles={loadingCardStyles()}>
          <Spinner size={SpinnerSize.large} label="A sincronizar componentes" />
          <Text variant="small" styles={{ root: { color: '#605e5c', textAlign: 'center' } }}>
            Pré-visualização 1 de 3 — indicador centrado
          </Text>
        </Stack>
      )}
      {idx === 1 && (
        <Stack tokens={{ childrenGap: 12 }} styles={loadingCardStyles()}>
          <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
            Pré-visualização 2 de 3 — esqueleto (shimmer)
          </Text>
          <Shimmer
            shimmerElements={[{ type: ShimmerElementType.line, height: 10, width: '55%' }]}
          />
          <Shimmer
            shimmerElements={[
              { type: ShimmerElementType.line, height: 32, width: '42%' },
              { type: ShimmerElementType.gap, width: 12, height: 32 },
              { type: ShimmerElementType.line, height: 32, width: '42%' },
            ]}
          />
          <Shimmer
            shimmerElements={[{ type: ShimmerElementType.line, height: 72, width: '100%' }]}
          />
        </Stack>
      )}
      {idx === 2 && (
        <Stack tokens={{ childrenGap: 12 }} styles={loadingCardStyles()}>
          <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
            Pré-visualização 3 de 3 — progresso e faixas
          </Text>
          <ProgressIndicator label="A carregar opções de interface" />
          <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 6, width: '28%' }]} />
          <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 6, width: '72%' }]} />
          <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 6, width: '48%' }]} />
        </Stack>
      )}
      <Text variant="small" styles={{ root: { color: '#a19f9d', fontStyle: 'italic' } }}>
        As três pré-visualizações alternam automaticamente enquanto espera.
      </Text>
    </Stack>
  );
}

export interface IFormManagerComponentsTabContentProps {
  loading: boolean;
  stepLayout: TFormStepLayoutKind;
  onStepLayoutChange: (v: TFormStepLayoutKind) => void;
  stepAccentPaletteSlot: TFormCustomButtonPaletteSlot | undefined;
  onStepAccentPaletteSlotChange: (v: TFormCustomButtonPaletteSlot) => void;
  stepNavButtons: TFormStepNavButtonsKind;
  onStepNavButtonsChange: (v: TFormStepNavButtonsKind) => void;
  formDataLoadingKind: TFormDataLoadingUiKind;
  onFormDataLoadingKindChange: (v: TFormDataLoadingUiKind) => void;
  defaultSubmitLoadingKind: TFormSubmitLoadingUiKind;
  onDefaultSubmitLoadingKindChange: (v: TFormSubmitLoadingUiKind) => void;
  historyEnabled: boolean;
  onHistoryEnabledChange: (enabled: boolean) => void;
  historyPresentationKind: TFormHistoryPresentationKind;
  onHistoryPresentationKindChange: (v: TFormHistoryPresentationKind) => void;
  historyButtonKind: TFormHistoryButtonKind;
  onHistoryButtonKindChange: (v: TFormHistoryButtonKind) => void;
  historyButtonLabel: string;
  onHistoryButtonLabelChange: (v: string) => void;
  historyButtonIcon: string;
  onHistoryButtonIconChange: (v: string) => void;
  historyPanelSubtitle: string;
  onHistoryPanelSubtitleChange: (v: string) => void;
  historyGroupTitles: string[];
  onHistoryGroupTitlesChange: (titles: string[]) => void;
  siteGroups: IGroupDetails[];
  siteGroupsSorted: IGroupDetails[];
  siteGroupsLoading: boolean;
  siteGroupsErr?: string;
  onRetryLoadSiteGroups: () => void;
  historyLayoutKind: TFormHistoryLayoutKind;
  onHistoryLayoutKindChange: (v: TFormHistoryLayoutKind) => void;
}

export function FormManagerComponentsTabContent(props: IFormManagerComponentsTabContentProps): JSX.Element {
  const [openSections, setOpenSections] = useState<Record<string, boolean>>({});
  const [historyGroupNameFilter, setHistoryGroupNameFilter] = useState('');

  const toggleSection = (id: string): void => {
    setOpenSections((prev) => ({ ...prev, [id]: !prev[id] }));
  };

  const isOpen = (id: string): boolean => openSections[id] === true;

  if (props.loading) {
    return <FormManagerComponentsLoadingLayouts />;
  }
  const theme = useTheme();
  const stepAccentPreview = useMemo(
    () => resolveStepUiAccentColor(theme, props.stepAccentPaletteSlot),
    [theme, props.stepAccentPaletteSlot]
  );
  const showIconFields = props.historyButtonKind === 'icon' || props.historyButtonKind === 'iconAndText';
  const showLabelField = props.historyButtonKind === 'text' || props.historyButtonKind === 'iconAndText';
  const showAriaOnlyLabel = props.historyButtonKind === 'icon';
  const historySiteGroupsFiltered = useMemo(
    () => filterSiteGroupsByNameQuery(props.siteGroupsSorted, historyGroupNameFilter),
    [props.siteGroupsSorted, historyGroupNameFilter]
  );
  return (
    <Stack tokens={{ childrenGap: 10 }}>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Expanda cada secção para configurar. Por defeito todas vêm fechadas.
      </Text>

      <FormManagerCollapseSection
        title="Carregar formulário / dados"
        isOpen={isOpen(SECTION_IDS.loadData)}
        onToggle={() => toggleSection(SECTION_IDS.loadData)}
      >
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Ao carregar campos da lista ou ao abrir um item para editar.
        </Text>
        <Dropdown
          label="Estilo de loading (dados)"
          options={FORM_DATA_LOADING_DROPDOWN_OPTIONS}
          selectedKey={props.formDataLoadingKind}
          onChange={(_, o) =>
            o && props.onFormDataLoadingKindChange(String(o.key) as TFormDataLoadingUiKind)
          }
        />
        <Stack
          horizontalAlign="center"
          styles={{
            root: {
              border: '1px solid #edebe9',
              borderRadius: 4,
              padding: 8,
              background: '#faf9f8',
              minHeight: 140,
            },
          }}
        >
          <FormDataLoadingView
            kind={props.formDataLoadingKind}
            message="Pré-visualização — carregar campos / item"
          />
        </Stack>
      </FormManagerCollapseSection>

      <FormManagerCollapseSection
        title="Gravar — loading ao gravar (padrão)"
        isOpen={isOpen(SECTION_IDS.submitLoading)}
        onToggle={() => toggleSection(SECTION_IDS.submitLoading)}
      >
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Usado em botões personalizados sem override. Cada botão pode definir outro estilo na aba
          Botões.
        </Text>
        <Dropdown
          label="Estilo de loading ao gravar (padrão)"
          options={FORM_SUBMIT_LOADING_DROPDOWN_OPTIONS}
          selectedKey={props.defaultSubmitLoadingKind}
          onChange={(_, o) =>
            o && props.onDefaultSubmitLoadingKindChange(String(o.key) as TFormSubmitLoadingUiKind)
          }
        />
      </FormManagerCollapseSection>

      <FormManagerCollapseSection
        title="Etapas — layout e navegação"
        isOpen={isOpen(SECTION_IDS.steps)}
        onToggle={() => toggleSection(SECTION_IDS.steps)}
      >
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Com mais do que uma etapa (aba Estrutura), o passador usa o layout escolhido. Os botões «anterior» /
          «próxima» no rodapé são independentes do passador.
        </Text>
        <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Layout das etapas no formulário</Text>
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Navegação visual entre etapas (quando existir mais de uma).
        </Text>
        <ThemePaletteSlotDropdown
          label="Cor de destaque (passador e botões de etapa)"
          selectedKey={props.stepAccentPaletteSlot ?? 'themePrimary'}
          onChange={(slot) => props.onStepAccentPaletteSlotChange(slot)}
        />
        <FormStepLayoutPicker
          value={props.stepLayout}
          onChange={props.onStepLayoutChange}
          accentColor={stepAccentPreview}
        />
        <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
          Botões «Etapa anterior» / «Próxima etapa»
        </Text>
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Estilo apenas dos botões de navegação no rodapé.
        </Text>
        <FormStepNavButtonsPicker
          value={props.stepNavButtons}
          onChange={props.onStepNavButtonsChange}
          accentColor={stepAccentPreview}
        />
      </FormManagerCollapseSection>

      <FormManagerCollapseSection
        title="Histórico de auditoria"
        isOpen={isOpen(SECTION_IDS.historyAudit)}
        onToggle={() => toggleSection(SECTION_IDS.historyAudit)}
      >
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Botão integrado no formulário, forma de abrir o histórico e aspeto. A lista SharePoint e a captação de logs
          configuram-se na aba «Lista de logs».
        </Text>
        <Toggle
          label="Ativar botão de histórico de auditoria"
          checked={props.historyEnabled}
          onChange={(_, c) => props.onHistoryEnabledChange(!!c)}
          onText="Ativo"
          offText="Inativo"
        />
        {props.historyEnabled && (
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { maxWidth: 560 } }}>
            <Dropdown
              label="Abrir histórico como"
              options={[
                { key: 'panel', text: 'Painel lateral' },
                { key: 'modal', text: 'Modal' },
                { key: 'collapse', text: 'Secção no formulário (abaixo dos botões)' },
              ]}
              selectedKey={props.historyPresentationKind}
              onChange={(_, o) =>
                o && props.onHistoryPresentationKindChange(String(o.key) as TFormHistoryPresentationKind)
              }
            />
            <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
              Aspeto do botão no formulário
            </Text>
            <ChoiceGroup
              options={HISTORY_BUTTON_KIND_OPTIONS}
              selectedKey={props.historyButtonKind}
              onChange={(_, o) =>
                o && props.onHistoryButtonKindChange(String(o.key) as TFormHistoryButtonKind)
              }
            />
            {showLabelField && (
              <TextField
                label="Texto do botão"
                value={props.historyButtonLabel}
                onChange={(_, v) => props.onHistoryButtonLabelChange(v ?? '')}
                placeholder="Histórico"
              />
            )}
            {showAriaOnlyLabel && (
              <TextField
                label="Nome acessível (tooltip / leitor de ecrã)"
                value={props.historyButtonLabel}
                onChange={(_, v) => props.onHistoryButtonLabelChange(v ?? '')}
                placeholder="Histórico"
              />
            )}
            {showIconFields && (
              <TextField
                label="Ícone Fluent (nome)"
                description="Ex.: History, Clock, TimelineProgress. Em «só ícone», use o subtítulo abaixo como tooltip."
                value={props.historyButtonIcon}
                onChange={(_, v) => props.onHistoryButtonIconChange(v ?? '')}
                placeholder="History"
              />
            )}
            <TextField
              label="Subtítulo / ajuda (painel de histórico e tooltip)"
              multiline
              rows={2}
              value={props.historyPanelSubtitle}
              onChange={(_, v) => props.onHistoryPanelSubtitleChange(v ?? '')}
            />
            <Text variant="small" styles={{ root: { fontWeight: 600, marginTop: 8 } }}>
              Grupos do SharePoint
            </Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Só utilizadores que pertençam a pelo menos um dos grupos marcados vêem o botão de histórico. Vazio =
              todos.
            </Text>
            <TextField
              placeholder="Filtrar grupos por nome"
              value={historyGroupNameFilter}
              onChange={(_: unknown, v?: string) => setHistoryGroupNameFilter(v ?? '')}
              styles={{ root: { maxWidth: 420 } }}
            />
            {props.siteGroupsLoading && <Spinner label="A carregar grupos do site…" />}
            {props.siteGroupsErr && (
              <>
                <MessageBar messageBarType={MessageBarType.warning}>{props.siteGroupsErr}</MessageBar>
                <DefaultButton text="Tentar carregar grupos novamente" onClick={props.onRetryLoadSiteGroups} />
              </>
            )}
            {!props.siteGroupsLoading ? (
              <Stack
                tokens={{ childrenGap: 6 }}
                styles={{
                  root: {
                    maxHeight: 240,
                    overflowY: 'auto',
                    border: '1px solid #edebe9',
                    borderRadius: 4,
                    padding: 8,
                    background: '#ffffff',
                  },
                }}
              >
                {(props.historyGroupTitles ?? [])
                  .filter(
                    (t) =>
                      !props.siteGroups.some((g) => normSpGroupTitle(g.Title) === normSpGroupTitle(t))
                  )
                  .filter((t) => {
                    const q = historyGroupNameFilter.trim().toLowerCase();
                    return !q || t.toLowerCase().includes(q);
                  })
                  .map((t, oi) => (
                    <Checkbox
                      key={`hist-orphan-grp-${oi}-${t}`}
                      label={`${t} (guardado; não na lista do site)`}
                      checked
                      onChange={(_, c) => {
                        if (c) return;
                        const cur = props.historyGroupTitles ?? [];
                        const n = normSpGroupTitle(t);
                        const next = cur.filter((x) => normSpGroupTitle(x) !== n);
                        props.onHistoryGroupTitlesChange(next);
                      }}
                    />
                  ))}
                {historySiteGroupsFiltered.map((g) => {
                  const cur = props.historyGroupTitles ?? [];
                  const n = normSpGroupTitle(g.Title);
                  const checked = cur.some((x) => normSpGroupTitle(x) === n);
                  return (
                    <Checkbox
                      key={g.Id}
                      label={g.Title}
                      title={g.Description || undefined}
                      checked={checked}
                      onChange={(_, c) => {
                        let next: string[];
                        if (c) {
                          next = checked ? cur : cur.concat([g.Title]);
                        } else {
                          next = cur.filter((x) => normSpGroupTitle(x) !== n);
                        }
                        props.onHistoryGroupTitlesChange(next);
                      }}
                    />
                  );
                })}
                {props.siteGroupsSorted.length > 0 &&
                  !historySiteGroupsFiltered.length &&
                  historyGroupNameFilter.trim() && (
                    <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                      Nenhum grupo corresponde ao filtro.
                    </Text>
                  )}
                {!props.siteGroupsSorted.length && !(props.historyGroupTitles ?? []).length && (
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Nenhum grupo no site.
                  </Text>
                )}
              </Stack>
            ) : null}
            <Text variant="small" styles={{ root: { fontWeight: 600, marginTop: 12 } }}>
              Estilo da lista de registos (dentro do painel)
            </Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Aplica-se ao conteúdo do histórico (registos da lista de log filtrados pelo lookup).
            </Text>
            <Dropdown
              label="Estilo da lista de registos"
              options={FORM_HISTORY_LAYOUT_DROPDOWN_OPTIONS}
              selectedKey={props.historyLayoutKind}
              onChange={(_, o) =>
                o && props.onHistoryLayoutKindChange(String(o.key) as TFormHistoryLayoutKind)
              }
            />
            <HistoryLayoutPreview kind={props.historyLayoutKind} />
          </Stack>
        )}
      </FormManagerCollapseSection>
    </Stack>
  );
}
