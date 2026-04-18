import * as React from 'react';
import { useState } from 'react';
import {
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  IconButton,
  Link,
  Icon,
  Panel,
  PanelType,
} from '@fluentui/react';
import type { TFormStepLayoutKind, TFormStepNavButtonsKind } from '../../core/config/types/formManager';
import { hexToRgbaString, STEP_UI_FALLBACK_ACCENT_HEX } from '../../core/formManager/formCustomButtonTheme';

export const FORM_STEP_LAYOUT_OPTIONS: {
  id: TFormStepLayoutKind;
  title: string;
  description: string;
}[] = [
  {
    id: 'rail',
    title: 'Trilho lateral',
    description: 'Números em coluna à esquerda, linha guia e títulos alinhados. Ideal para formulários longos.',
  },
  {
    id: 'segmented',
    title: 'Segmentos',
    description: 'Pílulas horizontais preenchidas; estilo familiar em apps empresariais.',
  },
  {
    id: 'timeline',
    title: 'Linha do tempo',
    description: 'Pontos sobre uma barra de progresso; mostra claramente o quanto falta.',
  },
  {
    id: 'cards',
    title: 'Cartões',
    description: 'Cada etapa em cartão com sombra; hierarquia visual forte e legível.',
  },
  {
    id: 'breadcrumb',
    title: 'Migalhas',
    description:
      'Trilho horizontal com separadores — como navegação; compacto em altura.',
  },
  {
    id: 'underline',
    title: 'Separadores (tabs)',
    description: 'Estilo de separadores com linha inferior no ativo; evoca separadores de browser ou CRM.',
  },
  {
    id: 'outline',
    title: 'Contorno',
    description: 'Etiquetas com contorno; fundo claro no ativo — visual limpo e legível.',
  },
  {
    id: 'compact',
    title: 'Compacto',
    description: 'Chips pequenos em fila; máximo de etapas visíveis sem ocupar altura.',
  },
  {
    id: 'steps',
    title: 'Passo numerado',
    description: 'Círculos numerados em linha com ligador; indica ordem e progresso.',
  },
  {
    id: 'minimal',
    title: 'Minimal',
    description: 'Texto inline com pontos médios; máxima leveza, ideal para poucas etapas.',
  },
];

export const FORM_STEP_LAYOUT_QUICK_IDS: TFormStepLayoutKind[] = [
  'rail',
  'segmented',
  'timeline',
  'cards',
  'breadcrumb',
];

export const FORM_STEP_LAYOUT_TOTAL = FORM_STEP_LAYOUT_OPTIONS.length;

const line = '#c8c6c4';
const muted = '#605e5c';
const done = '#107c10';

export interface IFormStepLayoutPickerProps {
  value: TFormStepLayoutKind;
  onChange: (id: TFormStepLayoutKind) => void;
  accentColor?: string;
}

export const FormStepLayoutMiniPreview: React.FC<{ kind: TFormStepLayoutKind; accentColor?: string }> = ({
  kind,
  accentColor,
}) => {
  const accent = accentColor ?? STEP_UI_FALLBACK_ACCENT_HEX;
  const w = 72;
  const h = 40;
  const dot = (activeDot: boolean, doneDot: boolean): React.ReactNode => (
    <div
      style={{
        width: 8,
        height: 8,
        borderRadius: '50%',
        background: doneDot ? done : activeDot ? accent : '#edebe9',
        flexShrink: 0,
      }}
    />
  );
  if (kind === 'rail') {
    return (
      <div style={{ width: w, height: h, display: 'flex', gap: 4 }}>
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2 }}>
          {dot(true, false)}
          <div style={{ width: 2, flex: 1, background: line, minHeight: 6 }} />
          {dot(false, false)}
        </div>
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', justifyContent: 'space-between', padding: '2px 0' }}>
          <div style={{ height: 4, background: accent, borderRadius: 2, opacity: 0.35 }} />
          <div style={{ height: 4, background: '#edebe9', borderRadius: 2 }} />
        </div>
      </div>
    );
  }
  if (kind === 'segmented') {
    return (
      <div style={{ width: w, height: h, display: 'flex', alignItems: 'center', gap: 3 }}>
        <div style={{ flex: 1, height: 18, borderRadius: 9, background: accent, opacity: 0.9 }} />
        <div style={{ flex: 1, height: 18, borderRadius: 9, background: '#edebe9' }} />
      </div>
    );
  }
  if (kind === 'timeline') {
    return (
      <div style={{ width: w, height: h, display: 'flex', flexDirection: 'column', justifyContent: 'center', gap: 4 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          {dot(true, false)}
          <div style={{ flex: 1, height: 3, margin: '0 4px', background: `linear-gradient(90deg, ${accent} 50%, ${line} 50%)`, borderRadius: 2 }} />
          {dot(false, false)}
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between' }}>
          <div style={{ width: 20, height: 3, background: '#edebe9', borderRadius: 2 }} />
          <div style={{ width: 20, height: 3, background: '#edebe9', borderRadius: 2 }} />
        </div>
      </div>
    );
  }
  if (kind === 'breadcrumb') {
    return (
      <div style={{ width: w, height: h, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 3 }}>
        <div style={{ width: 22, height: 8, borderRadius: 2, background: accent }} />
        <span style={{ color: line, fontSize: 11, lineHeight: 1, fontWeight: 700 }}>›</span>
        <div style={{ width: 16, height: 8, borderRadius: 2, background: '#edebe9' }} />
        <span style={{ color: line, fontSize: 11, lineHeight: 1, fontWeight: 700 }}>›</span>
        <div style={{ width: 16, height: 8, borderRadius: 2, background: '#edebe9' }} />
      </div>
    );
  }
  if (kind === 'cards') {
    return (
      <div style={{ width: w, height: h, display: 'flex', gap: 4, alignItems: 'stretch' }}>
        <div
          style={{
            flex: 1,
            borderRadius: 4,
            border: `2px solid ${accent}`,
            background: '#f3f9ff',
            boxShadow: '0 1px 4px rgba(0,0,0,0.08)',
          }}
        />
        <div style={{ flex: 1, borderRadius: 4, border: '1px solid #edebe9', background: '#fff' }} />
      </div>
    );
  }
  if (kind === 'underline') {
    return (
      <div
        style={{
          width: w,
          height: h,
          display: 'flex',
          flexDirection: 'column',
          justifyContent: 'flex-end',
          borderBottom: `2px solid ${line}`,
        }}
      >
        <div style={{ display: 'flex', height: 14, alignItems: 'flex-end', gap: 4 }}>
          <div style={{ flex: 1, height: 10, borderRadius: 2, background: '#edebe9' }} />
          <div
            style={{
              flex: 1,
              height: 10,
              borderRadius: 2,
              background: '#edebe9',
              borderBottom: `3px solid ${accent}`,
              marginBottom: -2,
            }}
          />
          <div style={{ flex: 1, height: 10, borderRadius: 2, background: '#edebe9' }} />
        </div>
      </div>
    );
  }
  if (kind === 'outline') {
    return (
      <div style={{ width: w, height: h, display: 'flex', alignItems: 'center', gap: 4 }}>
        <div
          style={{
            flex: 1,
            height: 22,
            borderRadius: 6,
            border: `2px solid ${accent}`,
            background: '#f3f9ff',
          }}
        />
        <div
          style={{ flex: 1, height: 22, borderRadius: 6, border: `1px solid ${line}`, background: '#fff' }}
        />
      </div>
    );
  }
  if (kind === 'compact') {
    return (
      <div style={{ width: w, height: h, display: 'flex', alignItems: 'center', gap: 3 }}>
        <div style={{ padding: '3px 8px', borderRadius: 4, background: accent, opacity: 0.9, height: 14 }} />
        <div style={{ padding: '3px 8px', borderRadius: 4, background: '#edebe9', height: 14, flex: 1, maxWidth: 28 }} />
        <div style={{ padding: '3px 8px', borderRadius: 4, background: '#edebe9', height: 14, flex: 1, maxWidth: 28 }} />
      </div>
    );
  }
  if (kind === 'steps') {
    return (
      <div style={{ width: w, height: h, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div
          style={{
            width: 14,
            height: 14,
            borderRadius: '50%',
            background: accent,
            color: '#fff',
            fontSize: 9,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontWeight: 700,
          }}
        >
          1
        </div>
        <div style={{ flex: 1, height: 2, background: line, margin: '0 2px' }} />
        <div
          style={{
            width: 14,
            height: 14,
            borderRadius: '50%',
            border: `2px solid ${line}`,
            background: '#fff',
            fontSize: 9,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontWeight: 700,
            color: muted,
          }}
        >
          2
        </div>
        <div style={{ flex: 1, height: 2, background: line, margin: '0 2px' }} />
        <div
          style={{
            width: 14,
            height: 14,
            borderRadius: '50%',
            border: `2px solid ${line}`,
            background: '#fff',
            fontSize: 9,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontWeight: 700,
            color: muted,
          }}
        >
          3
        </div>
      </div>
    );
  }
  if (kind === 'minimal') {
    return (
      <div
        style={{
          width: w,
          height: h,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          gap: 2,
          fontSize: 9,
          color: muted,
        }}
      >
        <span style={{ color: accent, fontWeight: 700 }}>A</span>
        <span>·</span>
        <span>B</span>
        <span>·</span>
        <span>C</span>
      </div>
    );
  }
  return (
    <div style={{ width: w, height: h, display: 'flex', gap: 4, alignItems: 'stretch' }}>
      <div style={{ flex: 1, borderRadius: 4, border: `2px solid ${accent}`, background: '#f3f9ff', boxShadow: '0 1px 4px rgba(0,0,0,0.08)' }} />
      <div style={{ flex: 1, borderRadius: 4, border: '1px solid #edebe9', background: '#fff' }} />
    </div>
  );
};

function layoutOptionButtonStyle(sel: boolean, accent: string): React.CSSProperties {
  return {
    display: 'flex',
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: 12,
    width: '100%',
    boxSizing: 'border-box',
    textAlign: 'left',
    font: 'inherit',
    padding: 14,
    borderRadius: 10,
    border: sel ? `2px solid ${accent}` : '2px solid #edebe9',
    background: sel ? '#f3f9ff' : '#fff',
    cursor: 'pointer',
    boxShadow: sel ? `0 4px 14px ${hexToRgbaString(accent, 0.12)}` : '0 1px 4px rgba(0,0,0,0.06)',
    transition: 'border 0.15s ease, box-shadow 0.15s ease',
  };
}

export const FormStepLayoutPicker: React.FC<IFormStepLayoutPickerProps> = ({ value, onChange, accentColor }) => {
  const acc = accentColor ?? STEP_UI_FALLBACK_ACCENT_HEX;
  const [galleryOpen, setGalleryOpen] = useState(false);
  const current = FORM_STEP_LAYOUT_OPTIONS.find((o) => o.id === value);

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Text variant="small" styles={{ root: { color: muted } }}>
        Sugestões rápidas — o layout escolhido aplica-se ao passador de etapas no formulário.
      </Text>
      <Stack horizontal wrap tokens={{ childrenGap: 10 }}>
        {FORM_STEP_LAYOUT_QUICK_IDS.map((id) => {
          const opt = FORM_STEP_LAYOUT_OPTIONS.find((o) => o.id === id);
          if (!opt) return null;
          const sel = value === opt.id;
          return (
            <button
              key={opt.id}
              type="button"
              onClick={() => onChange(opt.id)}
              style={{
                ...layoutOptionButtonStyle(sel, acc),
                maxWidth: 200,
                minWidth: 168,
                flex: '1 1 168px',
              }}
            >
              <span style={{ flexShrink: 0, lineHeight: 0 }}>
                <FormStepLayoutMiniPreview kind={opt.id} accentColor={acc} />
              </span>
              <span style={{ flex: 1, minWidth: 0, display: 'block' }}>
                <span
                  style={{
                    display: 'block',
                    fontWeight: 600,
                    color: '#323130',
                    fontSize: 13,
                    lineHeight: 1.3,
                  }}
                >
                  {opt.title}
                </span>
              </span>
            </button>
          );
        })}
      </Stack>
      {current && (
        <Text variant="small" styles={{ root: { color: '#323130' } }}>
          <span style={{ fontWeight: 600 }}>Atual:</span> {current.title}
        </Text>
      )}
      <Link onClick={() => setGalleryOpen(true)}>
        Ver mais — explorar os {FORM_STEP_LAYOUT_TOTAL} layouts
      </Link>

      <Panel
        isOpen={galleryOpen}
        type={PanelType.medium}
        headerText="Layouts de etapas"
        onDismiss={() => setGalleryOpen(false)}
        closeButtonAriaLabel="Fechar"
        isBlocking
        isFooterAtBottom
        onRenderFooterContent={() => (
          <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }} styles={{ root: { width: '100%' } }}>
            <DefaultButton text="Fechar" onClick={() => setGalleryOpen(false)} />
          </Stack>
        )}
        styles={{
          main: { maxWidth: 560, display: 'flex', flexDirection: 'column', maxHeight: '100%' },
          content: { flex: 1, display: 'flex', flexDirection: 'column', minHeight: 0, paddingBottom: 0 },
          footer: {
            flexShrink: 0,
            borderTop: '1px solid #edebe9',
            paddingTop: 16,
            paddingBottom: 16,
            background: '#faf9f8',
          },
        }}
      >
        <Stack
          tokens={{ childrenGap: 12 }}
          styles={{
            root: {
              paddingTop: 4,
              flex: 1,
              minHeight: 0,
              display: 'flex',
              flexDirection: 'column',
            },
          }}
        >
          <Text
            variant="small"
            styles={{
              root: { color: muted, lineHeight: 1.5, flexShrink: 0 },
            }}
          >
            Escolha um estilo para o indicador de etapas no topo do formulário. Pode voltar aqui a qualquer momento para
            alterar.
          </Text>
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: '1fr',
              gap: 12,
              flex: 1,
              minHeight: 0,
              overflowY: 'auto',
              overflowX: 'hidden',
              paddingRight: 4,
              WebkitOverflowScrolling: 'touch',
            }}
          >
            {FORM_STEP_LAYOUT_OPTIONS.map((opt) => {
              const sel = value === opt.id;
              return (
                <button
                  key={opt.id}
                  type="button"
                  onClick={() => {
                    onChange(opt.id);
                    setGalleryOpen(false);
                  }}
                  style={layoutOptionButtonStyle(sel, acc)}
                >
                  <span
                    style={{
                      flexShrink: 0,
                      lineHeight: 0,
                      transform: 'scale(1.12)',
                      transformOrigin: 'top left',
                    }}
                  >
                    <FormStepLayoutMiniPreview kind={opt.id} accentColor={acc} />
                  </span>
                  <span style={{ flex: 1, minWidth: 0, display: 'block' }}>
                    <span
                      style={{
                        display: 'block',
                        fontWeight: 600,
                        color: '#323130',
                        fontSize: 14,
                        lineHeight: 1.3,
                        marginBottom: 6,
                      }}
                    >
                      {opt.title}
                    </span>
                    <span
                      style={{
                        display: 'block',
                        color: muted,
                        fontSize: 12,
                        lineHeight: 1.45,
                        whiteSpace: 'normal',
                        wordBreak: 'break-word',
                      }}
                    >
                      {opt.description}
                    </span>
                  </span>
                </button>
              );
            })}
          </div>
        </Stack>
      </Panel>
    </Stack>
  );
};

export interface IFormStepNavigationProps {
  steps: { id: string; title: string }[];
  activeIndex: number;
  onStepSelect: (index: number) => void;
  layout: TFormStepLayoutKind;
  accentColor?: string;
}

export const FormStepNavigation: React.FC<IFormStepNavigationProps> = ({
  steps,
  activeIndex,
  onStepSelect,
  layout,
  accentColor,
}) => {
  if (steps.length <= 1) return null;
  const accent = accentColor ?? STEP_UI_FALLBACK_ACCENT_HEX;

  const go = (i: number): void => {
    if (i >= 0 && i < steps.length) onStepSelect(i);
  };

  if (layout === 'rail') {
    return (
      <div
        style={{
          borderRadius: 10,
          border: '1px solid #edebe9',
          padding: '16px 20px',
          background: '#faf9f8',
          boxShadow: '0 1px 3px rgba(0,0,0,0.06)',
        }}
      >
        {steps.map((st, i) => {
          const active = i === activeIndex;
          const completed = i < activeIndex;
          return (
            <div key={st.id} style={{ display: 'flex', minHeight: i === steps.length - 1 ? 'auto' : 52 }}>
              <div
                style={{
                  width: 44,
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  flexShrink: 0,
                }}
              >
                <button
                  type="button"
                  onClick={() => go(i)}
                  style={{
                    width: 36,
                    height: 36,
                    borderRadius: '50%',
                    border: 'none',
                    padding: 0,
                    cursor: 'pointer',
                    background: completed ? done : active ? accent : '#edebe9',
                    color: '#fff',
                    fontSize: 15,
                    fontWeight: 700,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    boxShadow: active ? `0 2px 8px ${hexToRgbaString(accent, 0.35)}` : 'none',
                  }}
                  aria-current={active ? 'step' : undefined}
                >
                  {completed ? '✓' : i + 1}
                </button>
                {i < steps.length - 1 && (
                  <div
                    style={{
                      width: 3,
                      flex: 1,
                      minHeight: 16,
                      marginTop: 4,
                      borderRadius: 2,
                      background: completed ? done : line,
                      opacity: completed ? 0.45 : 0.55,
                    }}
                  />
                )}
              </div>
              <div style={{ paddingLeft: 14, paddingBottom: i === steps.length - 1 ? 0 : 8, flex: 1 }}>
                <Text
                  styles={{
                    root: {
                      fontWeight: active ? 700 : completed ? 500 : 400,
                      color: active ? accent : completed ? '#323130' : muted,
                      fontSize: 15,
                      cursor: 'pointer',
                    },
                  }}
                  onClick={() => go(i)}
                >
                  {st.title}
                </Text>
                {active && (
                  <Text variant="small" styles={{ root: { color: muted, marginTop: 4, display: 'block' } }}>
                    Etapa {activeIndex + 1} de {steps.length}
                  </Text>
                )}
              </div>
            </div>
          );
        })}
      </div>
    );
  }

  if (layout === 'segmented') {
    return (
      <div
        style={{
          display: 'flex',
          flexWrap: 'wrap',
          gap: 8,
          padding: 6,
          borderRadius: 12,
          background: '#f3f2f1',
          border: '1px solid #edebe9',
        }}
      >
        {steps.map((st, i) => {
          const active = i === activeIndex;
          return (
            <button
              key={st.id}
              type="button"
              onClick={() => go(i)}
              style={{
                flex: '1 1 auto',
                minWidth: 100,
                padding: '10px 18px',
                borderRadius: 999,
                border: 'none',
                cursor: 'pointer',
                fontWeight: active ? 700 : 500,
                fontSize: 14,
                color: active ? '#fff' : '#323130',
                background: active ? accent : '#fff',
                boxShadow: active
                  ? `0 4px 12px ${hexToRgbaString(accent, 0.35)}`
                  : '0 1px 2px rgba(0,0,0,0.08)',
                transition: 'background 0.15s ease, color 0.15s ease',
              }}
            >
              {st.title}
            </button>
          );
        })}
      </div>
    );
  }

  if (layout === 'timeline') {
    return (
      <div style={{ padding: '8px 4px 4px' }}>
        <div style={{ position: 'relative', marginBottom: 28 }}>
          <div
            style={{
              position: 'absolute',
              left: 0,
              right: 0,
              top: 14,
              height: 4,
              borderRadius: 2,
              background: '#edebe9',
              zIndex: 0,
            }}
          />
          <div
            style={{
              position: 'absolute',
              left: 0,
              top: 14,
              height: 4,
              borderRadius: 2,
              background: `linear-gradient(90deg, ${accent}, ${accent})`,
              width: `${steps.length <= 1 ? 0 : (activeIndex / (steps.length - 1)) * 100}%`,
              zIndex: 1,
              transition: 'width 0.25s ease',
            }}
          />
          <div style={{ display: 'flex', justifyContent: 'space-between', position: 'relative', zIndex: 2 }}>
            {steps.map((st, i) => {
              const active = i === activeIndex;
              const done = i < activeIndex;
              return (
                <button
                  key={st.id}
                  type="button"
                  onClick={() => go(i)}
                  style={{
                    width: 28,
                    height: 28,
                    borderRadius: '50%',
                    border: active ? `3px solid ${accent}` : '3px solid #fff',
                    background: done ? accent : active ? '#fff' : '#edebe9',
                    cursor: 'pointer',
                    boxShadow: active
                      ? `0 0 0 2px ${hexToRgbaString(accent, 0.25)}`
                      : '0 1px 3px rgba(0,0,0,0.12)',
                    padding: 0,
                  }}
                  aria-label={st.title}
                />
              );
            })}
          </div>
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8 }}>
          {steps.map((st, i) => (
            <div key={st.id} style={{ flex: 1, textAlign: 'center', minWidth: 0 }}>
              <Text
                variant="small"
                styles={{
                  root: {
                    fontWeight: i === activeIndex ? 700 : 400,
                    color: i === activeIndex ? accent : muted,
                    cursor: 'pointer',
                    display: 'block',
                    lineHeight: 1.3,
                  },
                }}
                onClick={() => go(i)}
              >
                {st.title}
              </Text>
            </div>
          ))}
        </div>
      </div>
    );
  }

  if (layout === 'breadcrumb') {
    return (
      <nav
        aria-label="Etapas do formulário"
        style={{
          borderRadius: 10,
          border: '1px solid #e1dfdd',
          padding: '12px 16px',
          background: 'linear-gradient(180deg, #ffffff 0%, #f8f9fa 100%)',
          boxShadow: '0 1px 2px rgba(0,0,0,0.04)',
        }}
      >
        <div
          style={{
            display: 'flex',
            flexWrap: 'wrap',
            alignItems: 'center',
            gap: 2,
            rowGap: 6,
          }}
        >
          {steps.map((st, i) => {
            const active = i === activeIndex;
            const completed = i < activeIndex;
            return (
              <React.Fragment key={st.id}>
                {i > 0 && (
                  <Icon
                    iconName="ChevronRight"
                    styles={{
                      root: {
                        fontSize: 10,
                        color: line,
                        flexShrink: 0,
                        margin: '0 2px',
                        opacity: 0.85,
                      },
                    }}
                  />
                )}
                <button
                  type="button"
                  onClick={() => go(i)}
                  style={{
                    border: 'none',
                    cursor: 'pointer',
                    background: active ? hexToRgbaString(accent, 0.08) : 'transparent',
                    padding: '6px 12px',
                    borderRadius: 6,
                    maxWidth: '100%',
                    textAlign: 'left',
                    boxShadow: active ? `inset 0 -3px 0 0 ${accent}` : 'none',
                    transition: 'background 0.15s ease',
                  }}
                  aria-current={active ? 'step' : undefined}
                >
                  <span
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      gap: 6,
                      flexWrap: 'wrap',
                    }}
                  >
                    {completed && (
                      <Icon
                        iconName="CompletedSolid"
                        styles={{ root: { fontSize: 12, color: done, flexShrink: 0 } }}
                      />
                    )}
                    <Text
                      styles={{
                        root: {
                          fontWeight: active ? 700 : completed ? 600 : 400,
                          color: active ? accent : completed ? '#323130' : muted,
                          fontSize: 14,
                          lineHeight: 1.35,
                        },
                      }}
                    >
                      {st.title}
                    </Text>
                  </span>
                </button>
              </React.Fragment>
            );
          })}
        </div>
        <Text variant="small" styles={{ root: { color: muted, marginTop: 10, display: 'block' } }}>
          Etapa {activeIndex + 1} de {steps.length}
        </Text>
      </nav>
    );
  }

  if (layout === 'underline') {
    return (
      <div
        style={{
          borderBottom: `2px solid #edebe9`,
          borderRadius: '8px 8px 0 0',
        }}
      >
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 0 }}>
          {steps.map((st, i) => {
            const active = i === activeIndex;
            return (
              <button
                key={st.id}
                type="button"
                onClick={() => go(i)}
                style={{
                  padding: '12px 18px',
                  border: 'none',
                  background: 'transparent',
                  borderBottom: active ? `3px solid ${accent}` : '3px solid transparent',
                  marginBottom: -2,
                  color: active ? accent : muted,
                  fontWeight: active ? 700 : 500,
                  fontSize: 14,
                  cursor: 'pointer',
                  transition: 'color 0.15s ease',
                }}
                aria-current={active ? 'step' : undefined}
              >
                {st.title}
              </button>
            );
          })}
        </div>
      </div>
    );
  }

  if (layout === 'outline') {
    return (
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10, padding: '4px 2px' }}>
        {steps.map((st, i) => {
          const active = i === activeIndex;
          return (
            <button
              key={st.id}
              type="button"
              onClick={() => go(i)}
              style={{
                padding: '10px 16px',
                borderRadius: 8,
                border: active ? `2px solid ${accent}` : `2px solid ${line}`,
                background: active ? '#f3f9ff' : '#ffffff',
                color: active ? accent : '#323130',
                fontWeight: active ? 700 : 500,
                fontSize: 14,
                cursor: 'pointer',
                boxShadow: active ? `0 2px 8px ${hexToRgbaString(accent, 0.12)}` : 'none',
                transition: 'border 0.15s ease, background 0.15s ease',
              }}
              aria-current={active ? 'step' : undefined}
            >
              {st.title}
            </button>
          );
        })}
      </div>
    );
  }

  if (layout === 'compact') {
    return (
      <div
        style={{
          display: 'flex',
          flexWrap: 'wrap',
          alignItems: 'center',
          gap: 6,
          padding: '4px 0',
        }}
      >
        {steps.map((st, i) => {
          const active = i === activeIndex;
          return (
            <button
              key={st.id}
              type="button"
              onClick={() => go(i)}
              style={{
                padding: '4px 10px',
                borderRadius: 4,
                border: 'none',
                background: active ? accent : '#f3f2f1',
                color: active ? '#fff' : '#323130',
                fontWeight: active ? 700 : 500,
                fontSize: 12,
                cursor: 'pointer',
                boxShadow: active ? `0 2px 6px ${hexToRgbaString(accent, 0.3)}` : 'none',
              }}
              aria-current={active ? 'step' : undefined}
            >
              {st.title}
            </button>
          );
        })}
      </div>
    );
  }

  if (layout === 'steps') {
    return (
      <div style={{ padding: '8px 4px 12px' }}>
        <div style={{ display: 'flex', alignItems: 'center', width: '100%' }}>
          {steps.map((st, i) => {
            const active = i === activeIndex;
            const completed = i < activeIndex;
            const showLine = i < steps.length - 1;
            return (
              <React.Fragment key={st.id}>
                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', flex: '1 1 0', minWidth: 0 }}>
                  <button
                    type="button"
                    onClick={() => go(i)}
                    style={{
                      width: 32,
                      height: 32,
                      borderRadius: '50%',
                      border: active ? `3px solid ${accent}` : completed ? 'none' : `2px solid ${line}`,
                      background: completed ? done : active ? accent : '#fff',
                      color: completed || active ? '#fff' : muted,
                      fontSize: 13,
                      fontWeight: 700,
                      cursor: 'pointer',
                      padding: 0,
                      flexShrink: 0,
                      boxShadow: active
                        ? `0 2px 10px ${hexToRgbaString(accent, 0.35)}`
                        : '0 1px 3px rgba(0,0,0,0.08)',
                    }}
                    aria-current={active ? 'step' : undefined}
                    aria-label={st.title}
                  >
                    {completed ? '✓' : i + 1}
                  </button>
                  <Text
                    variant="small"
                    styles={{
                      root: {
                        marginTop: 8,
                        fontWeight: active ? 700 : 500,
                        color: active ? accent : muted,
                        textAlign: 'center',
                        cursor: 'pointer',
                        lineHeight: 1.3,
                        maxWidth: '100%',
                      },
                    }}
                    onClick={() => go(i)}
                  >
                    {st.title}
                  </Text>
                </div>
                {showLine && (
                  <div
                    style={{
                      flex: '0 0 12px',
                      height: 3,
                      alignSelf: 'flex-start',
                      marginTop: 14,
                      borderRadius: 2,
                      background: i < activeIndex ? done : line,
                      opacity: i < activeIndex ? 0.85 : 0.5,
                    }}
                  />
                )}
              </React.Fragment>
            );
          })}
        </div>
      </div>
    );
  }

  if (layout === 'minimal') {
    return (
      <div
        style={{
          display: 'flex',
          flexWrap: 'wrap',
          alignItems: 'center',
          padding: '6px 0',
          gap: 2,
        }}
      >
        {steps.map((st, i) => {
          const active = i === activeIndex;
          const completed = i < activeIndex;
          return (
            <React.Fragment key={st.id}>
              {i > 0 && (
                <span style={{ color: line, padding: '0 6px', userSelect: 'none', fontWeight: 300 }}>
                  ·
                </span>
              )}
              <button
                type="button"
                onClick={() => go(i)}
                style={{
                  border: 'none',
                  background: 'none',
                  padding: '4px 2px',
                  cursor: 'pointer',
                  color: active ? accent : completed ? '#323130' : muted,
                  fontWeight: active ? 700 : completed ? 500 : 400,
                  fontSize: 14,
                  textDecoration: active ? 'underline' : 'none',
                  textUnderlineOffset: 3,
                }}
                aria-current={active ? 'step' : undefined}
              >
                {st.title}
              </button>
            </React.Fragment>
          );
        })}
      </div>
    );
  }

  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12 }}>
      {steps.map((st, i) => {
        const active = i === activeIndex;
        return (
          <button
            key={st.id}
            type="button"
            onClick={() => go(i)}
            style={{
              flex: '1 1 140px',
              maxWidth: 220,
              padding: '14px 16px',
              borderRadius: 10,
              border: active ? `2px solid ${accent}` : '1px solid #edebe9',
              background: active ? '#f3f9ff' : '#fff',
              cursor: 'pointer',
              textAlign: 'left',
              boxShadow: active
                ? `0 6px 20px ${hexToRgbaString(accent, 0.15)}`
                : '0 2px 8px rgba(0,0,0,0.06)',
              transition: 'box-shadow 0.15s ease, border 0.15s ease',
            }}
          >
            <Text variant="small" styles={{ root: { color: muted, fontWeight: 600, marginBottom: 6, display: 'block' } }}>
              Etapa {i + 1}
            </Text>
            <Text styles={{ root: { fontWeight: active ? 700 : 500, color: active ? accent : '#323130', fontSize: 14 } }}>
              {st.title}
            </Text>
          </button>
        );
      })}
    </div>
  );
};

export const FORM_STEP_NAV_BUTTONS_OPTIONS: {
  id: TFormStepNavButtonsKind;
  title: string;
  description: string;
}[] = [
  {
    id: 'fluent',
    title: 'Fluent padrão',
    description: 'Secundário «Etapa anterior» e primário «Próxima etapa», como botões Fluent.',
  },
  {
    id: 'pills',
    title: 'Pílulas',
    description: 'Dois botões largos em formato de cápsula.',
  },
  {
    id: 'dots',
    title: 'Bolinhas e setas',
    description: 'Indicadores da etapa atual e botões circulares com setas.',
  },
  {
    id: 'icons',
    title: 'Só ícones',
    description: 'Setas compactas; o significado fica nos tooltips.',
  },
  {
    id: 'links',
    title: 'Ligações de texto',
    description: 'Estilo de hiperligação, visual leve.',
  },
  {
    id: 'split',
    title: 'Extremos',
    description: '«Anterior» à esquerda e «Próxima» à direita — ocupa toda a largura do rodapé.',
  },
  {
    id: 'stacked',
    title: 'Empilhado',
    description: 'Dois botões em coluna, largura total — bom em telemóvel ou formulários estreitos.',
  },
  {
    id: 'ghost',
    title: 'Contorno',
    description: 'Fundo transparente com contorno; próxima etapa com realce em cor.',
  },
  {
    id: 'toolbar',
    title: 'Barra cinza',
    description: 'Faixa tipo barra de ferramentas, botões agrupados ao centro.',
  },
  {
    id: 'compact',
    title: 'Compacto',
    description: 'Botões mais baixos e menos padding — máximo de conteúdo visível.',
  },
];

export const FORM_STEP_NAV_BUTTONS_QUICK_IDS: TFormStepNavButtonsKind[] = [
  'fluent',
  'pills',
  'dots',
  'icons',
  'links',
];

export const FORM_STEP_NAV_BUTTONS_TOTAL = FORM_STEP_NAV_BUTTONS_OPTIONS.length;

export const FormStepNavButtonsMiniPreview: React.FC<{ kind: TFormStepNavButtonsKind; accentColor?: string }> = ({
  kind,
  accentColor,
}) => {
  const accent = accentColor ?? STEP_UI_FALLBACK_ACCENT_HEX;
  const w = 72;
  const pill = (filled: boolean): React.ReactNode => (
    <div
      style={{
        height: 14,
        borderRadius: 999,
        flex: 1,
        background: filled ? accent : '#edebe9',
        opacity: filled ? 1 : 0.85,
      }}
    />
  );
  if (kind === 'fluent') {
    return (
      <div style={{ width: w, height: 32, display: 'flex', gap: 4, alignItems: 'center' }}>
        <div style={{ flex: 1, height: 14, borderRadius: 2, border: `1px solid ${line}`, background: '#fff' }} />
        <div style={{ flex: 1, height: 14, borderRadius: 2, background: accent }} />
      </div>
    );
  }
  if (kind === 'pills') {
    return (
      <div style={{ width: w, height: 32, display: 'flex', gap: 5, alignItems: 'center' }}>
        {pill(false)}
        {pill(true)}
      </div>
    );
  }
  if (kind === 'dots') {
    return (
      <div style={{ width: w, height: 32, display: 'flex', alignItems: 'center', gap: 4, justifyContent: 'center' }}>
        <div style={{ width: 12, height: 12, borderRadius: '50%', border: `1px solid ${line}` }} />
        <div style={{ display: 'flex', gap: 3 }}>
          {[0, 1, 2].map((i) => (
            <div
              key={i}
              style={{
                width: 6,
                height: 6,
                borderRadius: '50%',
                background: i === 1 ? accent : '#edebe9',
              }}
            />
          ))}
        </div>
        <div style={{ width: 12, height: 12, borderRadius: '50%', border: `1px solid ${line}` }} />
      </div>
    );
  }
  if (kind === 'icons') {
    return (
      <div style={{ width: w, height: 32, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10 }}>
        <span style={{ color: muted, fontSize: 14 }}>‹</span>
        <span style={{ color: accent, fontSize: 14 }}>›</span>
      </div>
    );
  }
  if (kind === 'links') {
    return (
      <div style={{ width: w, height: 32, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div style={{ height: 2, width: 22, background: accent, borderRadius: 1 }} />
        <div style={{ height: 2, width: 22, background: accent, borderRadius: 1 }} />
      </div>
    );
  }
  if (kind === 'split') {
    return (
      <div style={{ width: w, height: 32, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div style={{ width: 28, height: 14, borderRadius: 2, border: `1px solid ${line}`, background: '#fff' }} />
        <div style={{ width: 28, height: 14, borderRadius: 2, background: accent }} />
      </div>
    );
  }
  if (kind === 'stacked') {
    return (
      <div style={{ width: w, height: 32, display: 'flex', flexDirection: 'column', gap: 4, justifyContent: 'center' }}>
        <div style={{ height: 11, borderRadius: 2, border: `1px solid ${line}`, background: '#fff' }} />
        <div style={{ height: 11, borderRadius: 2, background: accent }} />
      </div>
    );
  }
  if (kind === 'ghost') {
    return (
      <div style={{ width: w, height: 32, display: 'flex', gap: 5, alignItems: 'center' }}>
        <div style={{ flex: 1, height: 14, borderRadius: 4, border: `1px solid ${line}`, background: 'transparent' }} />
        <div style={{ flex: 1, height: 14, borderRadius: 4, border: `1px solid ${accent}`, background: 'transparent' }} />
      </div>
    );
  }
  if (kind === 'toolbar') {
    return (
      <div
        style={{
          width: w,
          height: 32,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          gap: 6,
          background: '#f3f2f1',
          borderRadius: 6,
          border: '1px solid #edebe9',
        }}
      >
        <div style={{ width: 26, height: 12, borderRadius: 2, border: `1px solid ${line}`, background: '#fff' }} />
        <div style={{ width: 26, height: 12, borderRadius: 2, background: accent }} />
      </div>
    );
  }
  if (kind === 'compact') {
    return (
      <div style={{ width: w, height: 32, display: 'flex', gap: 4, alignItems: 'center' }}>
        <div style={{ flex: 1, height: 11, borderRadius: 2, border: `1px solid ${line}`, background: '#fff' }} />
        <div style={{ flex: 1, height: 11, borderRadius: 2, background: accent }} />
      </div>
    );
  }
  return (
    <div style={{ width: w, height: 32, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
      <div style={{ height: 2, width: 22, background: accent, borderRadius: 1 }} />
      <div style={{ height: 2, width: 22, background: accent, borderRadius: 1 }} />
    </div>
  );
};

export interface IFormStepNavButtonsPickerProps {
  value: TFormStepNavButtonsKind;
  onChange: (id: TFormStepNavButtonsKind) => void;
  accentColor?: string;
}

export const FormStepNavButtonsPicker: React.FC<IFormStepNavButtonsPickerProps> = ({
  value,
  onChange,
  accentColor,
}) => {
  const acc = accentColor ?? STEP_UI_FALLBACK_ACCENT_HEX;
  const [galleryOpen, setGalleryOpen] = useState(false);
  const current = FORM_STEP_NAV_BUTTONS_OPTIONS.find((o) => o.id === value);

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Text variant="small" styles={{ root: { color: muted } }}>
        Botões no rodapé do formulário (com mais do que uma etapa). Sugestões rápidas:
      </Text>
      <Stack horizontal wrap tokens={{ childrenGap: 10 }}>
        {FORM_STEP_NAV_BUTTONS_QUICK_IDS.map((id) => {
          const opt = FORM_STEP_NAV_BUTTONS_OPTIONS.find((o) => o.id === id);
          if (!opt) return null;
          const sel = value === opt.id;
          return (
            <button
              key={opt.id}
              type="button"
              onClick={() => onChange(opt.id)}
              style={{
                ...layoutOptionButtonStyle(sel, acc),
                maxWidth: 200,
                minWidth: 160,
                flex: '1 1 160px',
              }}
            >
              <span style={{ flexShrink: 0, lineHeight: 0 }}>
                <FormStepNavButtonsMiniPreview kind={opt.id} accentColor={acc} />
              </span>
              <span style={{ flex: 1, minWidth: 0, display: 'block' }}>
                <span
                  style={{
                    display: 'block',
                    fontWeight: 600,
                    color: '#323130',
                    fontSize: 13,
                    lineHeight: 1.3,
                  }}
                >
                  {opt.title}
                </span>
              </span>
            </button>
          );
        })}
      </Stack>
      {current && (
        <Text variant="small" styles={{ root: { color: '#323130' } }}>
          <span style={{ fontWeight: 600 }}>Atual:</span> {current.title}
        </Text>
      )}
      <Link onClick={() => setGalleryOpen(true)}>
        Ver mais — explorar os {FORM_STEP_NAV_BUTTONS_TOTAL} estilos de botões
      </Link>

      <Panel
        isOpen={galleryOpen}
        type={PanelType.medium}
        headerText="Botões «Etapa anterior» / «Próxima etapa»"
        onDismiss={() => setGalleryOpen(false)}
        closeButtonAriaLabel="Fechar"
        isBlocking
        isFooterAtBottom
        onRenderFooterContent={() => (
          <Stack horizontal horizontalAlign="end" styles={{ root: { width: '100%' } }}>
            <DefaultButton text="Fechar" onClick={() => setGalleryOpen(false)} />
          </Stack>
        )}
        styles={{
          main: { maxWidth: 560, display: 'flex', flexDirection: 'column', maxHeight: '100%' },
          content: { flex: 1, display: 'flex', flexDirection: 'column', minHeight: 0, paddingBottom: 0 },
          footer: {
            flexShrink: 0,
            borderTop: '1px solid #edebe9',
            paddingTop: 16,
            paddingBottom: 16,
            background: '#faf9f8',
          },
        }}
      >
        <Stack
          tokens={{ childrenGap: 12 }}
          styles={{ root: { paddingTop: 4, flex: 1, minHeight: 0, display: 'flex', flexDirection: 'column' } }}
        >
          <Text variant="small" styles={{ root: { color: muted, lineHeight: 1.5, flexShrink: 0 } }}>
            Estes estilos aplicam-se apenas aos dois botões de navegação no rodapé (não ao passador de etapas em cima).
          </Text>
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: '1fr',
              gap: 12,
              flex: 1,
              minHeight: 0,
              overflowY: 'auto',
              overflowX: 'hidden',
              paddingRight: 4,
              WebkitOverflowScrolling: 'touch',
            }}
          >
            {FORM_STEP_NAV_BUTTONS_OPTIONS.map((opt) => {
              const sel = value === opt.id;
              return (
                <button
                  key={opt.id}
                  type="button"
                  onClick={() => {
                    onChange(opt.id);
                    setGalleryOpen(false);
                  }}
                  style={layoutOptionButtonStyle(sel, acc)}
                >
                  <span
                    style={{
                      flexShrink: 0,
                      lineHeight: 0,
                      transform: 'scale(1.12)',
                      transformOrigin: 'top left',
                    }}
                  >
                    <FormStepNavButtonsMiniPreview kind={opt.id} accentColor={acc} />
                  </span>
                  <span style={{ flex: 1, minWidth: 0, display: 'block' }}>
                    <span
                      style={{
                        display: 'block',
                        fontWeight: 600,
                        color: '#323130',
                        fontSize: 14,
                        lineHeight: 1.3,
                        marginBottom: 6,
                      }}
                    >
                      {opt.title}
                    </span>
                    <span
                      style={{
                        display: 'block',
                        color: muted,
                        fontSize: 12,
                        lineHeight: 1.45,
                        whiteSpace: 'normal',
                        wordBreak: 'break-word',
                      }}
                    >
                      {opt.description}
                    </span>
                  </span>
                </button>
              );
            })}
          </div>
        </Stack>
      </Panel>
    </Stack>
  );
};

export interface IFormStepPrevNextNavProps {
  variant: TFormStepNavButtonsKind;
  stepIndex: number;
  stepCount: number;
  onPrev: () => void;
  onNext: () => void;
  disabled?: boolean;
  accentColor?: string;
}

export const FormStepPrevNextNav: React.FC<IFormStepPrevNextNavProps> = ({
  variant,
  stepIndex,
  stepCount,
  onPrev,
  onNext,
  disabled,
  accentColor,
}) => {
  if (stepCount <= 1) return null;
  const accent = accentColor ?? STEP_UI_FALLBACK_ACCENT_HEX;
  const canPrev = stepIndex > 0;
  const canNext = stepIndex < stepCount - 1;
  const prevLabel = 'Etapa anterior';
  const nextLabel = 'Próxima etapa';

  const pillBtn = (primary: boolean, label: string, onClick: () => void, can: boolean): React.ReactNode => (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled || !can}
      style={{
        padding: '10px 20px',
        borderRadius: 999,
        border: 'none',
        cursor: disabled || !can ? 'default' : 'pointer',
        fontWeight: primary ? 700 : 600,
        fontSize: 14,
        color: primary ? '#fff' : '#323130',
        background: primary ? accent : '#fff',
        boxShadow: primary ? `0 4px 12px ${hexToRgbaString(accent, 0.35)}` : '0 1px 3px rgba(0,0,0,0.1)',
        opacity: can ? 1 : 0.45,
      }}
    >
      {label}
    </button>
  );

  if (variant === 'dots') {
    return (
      <div
        style={{
          flexBasis: '100%',
          width: '100%',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          gap: 12,
          flexWrap: 'wrap',
          paddingTop: 4,
        }}
      >
        <button
          type="button"
          aria-label={prevLabel}
          onClick={onPrev}
          disabled={disabled || !canPrev}
          style={{
            width: 36,
            height: 36,
            borderRadius: '50%',
            border: `1px solid ${line}`,
            background: '#fff',
            cursor: disabled || !canPrev ? 'default' : 'pointer',
            fontSize: 18,
            color: canPrev ? accent : muted,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: 0,
            opacity: canPrev ? 1 : 0.45,
          }}
        >
          ‹
        </button>
        <div style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap', justifyContent: 'center' }}>
          {(() => {
            const nodes: React.ReactNode[] = [];
            for (let i = 0; i < stepCount; i++) {
              nodes.push(
                <div
                  key={i}
                  style={{
                    width: i === stepIndex ? 10 : 8,
                    height: i === stepIndex ? 10 : 8,
                    borderRadius: '50%',
                    background: i === stepIndex ? accent : i < stepIndex ? done : '#edebe9',
                    boxShadow:
                      i === stepIndex ? `0 0 0 2px ${hexToRgbaString(accent, 0.25)}` : undefined,
                  }}
                />
              );
            }
            return nodes;
          })()}
        </div>
        <button
          type="button"
          aria-label={nextLabel}
          onClick={onNext}
          disabled={disabled || !canNext}
          style={{
            width: 36,
            height: 36,
            borderRadius: '50%',
            border: `1px solid ${line}`,
            background: '#fff',
            cursor: disabled || !canNext ? 'default' : 'pointer',
            fontSize: 18,
            color: canNext ? accent : muted,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: 0,
            opacity: canNext ? 1 : 0.45,
          }}
        >
          ›
        </button>
      </div>
    );
  }

  if (variant === 'pills') {
    return (
      <>
        {canPrev && pillBtn(false, prevLabel, onPrev, canPrev)}
        {canNext && pillBtn(true, nextLabel, onNext, canNext)}
      </>
    );
  }

  if (variant === 'icons') {
    return (
      <>
        {canPrev && (
          <IconButton
            iconProps={{ iconName: 'ChevronLeft' }}
            title={prevLabel}
            ariaLabel={prevLabel}
            onClick={onPrev}
            disabled={disabled || !canPrev}
          />
        )}
        {canNext && (
          <IconButton
            iconProps={{ iconName: 'ChevronRight' }}
            title={nextLabel}
            ariaLabel={nextLabel}
            onClick={onNext}
            disabled={disabled || !canNext}
          />
        )}
      </>
    );
  }

  if (variant === 'links') {
    return (
      <>
        {canPrev && (
          <Link
            href="#"
            onClick={(e) => {
              e.preventDefault();
              if (!disabled && canPrev) onPrev();
            }}
            styles={{ root: { cursor: disabled || !canPrev ? 'default' : 'pointer', opacity: canPrev ? 1 : 0.45 } }}
          >
            ← {prevLabel}
          </Link>
        )}
        {canNext && (
          <Link
            href="#"
            onClick={(e) => {
              e.preventDefault();
              if (!disabled && canNext) onNext();
            }}
            styles={{ root: { cursor: disabled || !canNext ? 'default' : 'pointer', opacity: canNext ? 1 : 0.45 } }}
          >
            {nextLabel} →
          </Link>
        )}
      </>
    );
  }

  if (variant === 'split') {
    return (
      <div
        style={{
          flexBasis: '100%',
          width: '100%',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          gap: 12,
          flexWrap: 'wrap',
        }}
      >
        <div>{canPrev && <DefaultButton text={prevLabel} onClick={onPrev} disabled={disabled || !canPrev} />}</div>
        <div>{canNext && <PrimaryButton text={nextLabel} onClick={onNext} disabled={disabled || !canNext} />}</div>
      </div>
    );
  }

  if (variant === 'stacked') {
    return (
      <Stack styles={{ root: { width: '100%', flexBasis: '100%' } }} tokens={{ childrenGap: 8 }}>
        {canPrev && (
          <DefaultButton
            text={prevLabel}
            onClick={onPrev}
            disabled={disabled || !canPrev}
            styles={{ root: { width: '100%' } }}
          />
        )}
        {canNext && (
          <PrimaryButton
            text={nextLabel}
            onClick={onNext}
            disabled={disabled || !canNext}
            styles={{ root: { width: '100%' } }}
          />
        )}
      </Stack>
    );
  }

  if (variant === 'ghost') {
    return (
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
        {canPrev && (
          <button
            type="button"
            onClick={onPrev}
            disabled={disabled || !canPrev}
            style={{
              padding: '10px 18px',
              borderRadius: 4,
              border: `1px solid ${line}`,
              background: 'transparent',
              cursor: disabled || !canPrev ? 'default' : 'pointer',
              fontWeight: 600,
              fontSize: 14,
              color: '#323130',
              opacity: canPrev ? 1 : 0.45,
              fontFamily: 'inherit',
            }}
          >
            {prevLabel}
          </button>
        )}
        {canNext && (
          <button
            type="button"
            onClick={onNext}
            disabled={disabled || !canNext}
            style={{
              padding: '10px 18px',
              borderRadius: 4,
              border: `2px solid ${accent}`,
              background: 'transparent',
              cursor: disabled || !canNext ? 'default' : 'pointer',
              fontWeight: 700,
              fontSize: 14,
              color: accent,
              opacity: canNext ? 1 : 0.45,
              fontFamily: 'inherit',
            }}
          >
            {nextLabel}
          </button>
        )}
      </div>
    );
  }

  if (variant === 'toolbar') {
    return (
      <div
        style={{
          flexBasis: '100%',
          width: '100%',
          display: 'flex',
          justifyContent: 'center',
          flexWrap: 'wrap',
          gap: 10,
          padding: '10px 14px',
          background: '#f3f2f1',
          borderRadius: 8,
          border: '1px solid #edebe9',
        }}
      >
        {canPrev && <DefaultButton text={prevLabel} onClick={onPrev} disabled={disabled || !canPrev} />}
        {canNext && <PrimaryButton text={nextLabel} onClick={onNext} disabled={disabled || !canNext} />}
      </div>
    );
  }

  if (variant === 'compact') {
    return (
      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
        {canPrev && (
          <DefaultButton
            text={prevLabel}
            onClick={onPrev}
            disabled={disabled || !canPrev}
            styles={{
              root: {
                minHeight: 28,
                paddingLeft: 12,
                paddingRight: 12,
                paddingTop: 4,
                paddingBottom: 4,
                fontSize: 13,
              },
            }}
          />
        )}
        {canNext && (
          <PrimaryButton
            text={nextLabel}
            onClick={onNext}
            disabled={disabled || !canNext}
            styles={{
              root: {
                minHeight: 28,
                paddingLeft: 12,
                paddingRight: 12,
                paddingTop: 4,
                paddingBottom: 4,
                fontSize: 13,
              },
            }}
          />
        )}
      </div>
    );
  }

  return (
    <>
      {canPrev && <DefaultButton text={prevLabel} onClick={onPrev} disabled={disabled || !canPrev} />}
      {canNext && <PrimaryButton text={nextLabel} onClick={onNext} disabled={disabled || !canNext} />}
    </>
  );
};
